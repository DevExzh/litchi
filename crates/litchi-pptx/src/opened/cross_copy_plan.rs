//! Bounded cross-presentation slide-copy planning.
//!
//! A cross-package copy is deliberately more restrictive than a same-package
//! copy.  The source slide's private owned-role closure is copied byte-for-byte
//! into the destination, while the source layout is never copied: it is reused
//! only after the layout/master/theme inheritance graph has been proven
//! equivalent to the destination slide's selected layout.  The resulting
//! destination graph is captured as an exact, complete-revision-bound patch.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet, TryReserveError};
use std::fmt;
use std::io::{self, Write};
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, CompressedPartTransfer, OpcPackage, PackURI, Part, TargetMode};
use sha2::{Digest, Sha256};

use super::copy_plan::{
    available_copy_name_avoiding, collect_owned_closure, has_signature_infrastructure, reject_mce,
    reject_unknown_non_part_members, resolve_slide, validate_registered_layout,
    validate_slide_surface,
};
use super::model::{Limits, Slide, Snapshot, invalid};
use super::patch::Patch;
use crate::{Error, Result, SlideCopyRefusal};

/// Current durable cross-slide-copy header magic.
///
/// Bumped from `LPCP0003` by change 0742: a candidate may carry eligible copied
/// image members' verified source-compressed bytes instead of deflating them
/// again, which changes the serialized-archive revisions this header embeds,
/// and the header now records which of the two encodings the candidate used.
/// `LPCP0003` was itself bumped from `LPCP0002` by change 0655, when three of
/// the six embedded revisions became `litchi-pptx-opened-v2` complete-package
/// revisions; the physical revisions remain `litchi-pptx-cross-physical-v2`.
const PATCH_FORMAT: crate::DurablePatchFormat = crate::DurablePatchFormat::CrossSlideCopyV4;
const PATCH_MAGIC: &[u8; 8] = PATCH_FORMAT.magic();
/// Recognized but superseded durable cross-slide-copy formats, refused by
/// name before any header field is read.
const SUPERSEDED_PATCH_FORMATS: [crate::DurablePatchFormat; 2] = [
    crate::DurablePatchFormat::CrossSlideCopyV2,
    crate::DurablePatchFormat::CrossSlideCopyV3,
];
const PATCH_HEADER_BYTES: usize = PATCH_MAGIC.len() + (6 * 32) + (4 * 4) + 8 + 4 + 1 + 8;
const TRANSITIONAL_PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const TRANSITIONAL_DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_CHART: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const TRANSITIONAL_DIAGRAM: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const STRICT_DIAGRAM: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const TRANSITIONAL_CHART_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
const STRICT_CHART_DRAWING: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chartDrawing";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TRANSITIONAL_REL_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// How a cross-presentation candidate archive encodes its copied image
/// members.
///
/// The encoding is decided from the source's bytes when a copy is first
/// planned and recorded in the plan and in its durable patch, because it
/// determines the serialized-archive revisions both carry. Every later proof
/// of the same copy — application, durable-patch application in either
/// direction — rebuilds the candidate under the recorded encoding: a
/// recorded recompressed encoding transfers nothing, and a recorded
/// source-compressed one transfers what the source's bytes lend, which for
/// the recorded source revision is what planning transferred.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CopiedMedia {
    /// Every copied member is deflated again from its decoded bytes: the only
    /// encoding before change 0742, and the encoding of every copy whose
    /// source lends no eligible image.
    Recompressed,
    /// At least one copied member is an eligible image whose verified
    /// source-compressed bytes were framed verbatim; every other copied member
    /// is deflated again.
    SourceCompressed,
}

impl CopiedMedia {
    const fn wire(self) -> u8 {
        match self {
            Self::Recompressed => 0,
            Self::SourceCompressed => 1,
        }
    }

    fn from_wire(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Recompressed),
            1 => Ok(Self::SourceCompressed),
            _ => Err(invalid(
                "cross-slide durable patch has an unknown copied-media encoding",
            )),
        }
    }
}

/// Which copied-media encoding one candidate build may use.
#[derive(Clone, Copy)]
enum MediaPolicy {
    /// First planning: transfer every copied image the bytes the source
    /// publishes lend. The decision depends on those bytes alone, never on
    /// either package's edit history or allocation identity.
    Classify,
    /// Re-proving a plan or durable patch: use the encoding it recorded.
    Recorded(CopiedMedia),
}

/// The copied members a candidate frames from source-compressed bytes.
///
/// Built in two steps, so that no capture is taken before every limit has
/// been checked: [`classify_media`] applies the format rule and the source's
/// header-only eligibility, and [`MediaTransfers::capture`] verifies each
/// eligible member.
struct MediaTransfers<'s> {
    /// The package the captures come from: the source itself when it is an
    /// unmodified owned source, otherwise the reopen of its serialization.
    source_view: Cow<'s, OpcPackage>,
    /// Indexes into the planned parts, ascending, with their captures once
    /// taken. A capture is `None` before [`MediaTransfers::capture`] and in a
    /// release build that reuses a retained archive built from exactly these
    /// members.
    members: Vec<(usize, Option<CompressedPartTransfer>)>,
    /// Declared compressed bytes of the eligible members, which their captures
    /// hold while the candidate is built.
    capture_bytes: usize,
}

/// What one candidate build is given besides the two snapshots.
///
/// Carried as one argument because [`build_candidate`] already takes the
/// crate's maximum and `clippy::too_many_arguments` is deny.
struct CandidateInputs<'a> {
    archive: CandidateArchive<'a>,
    /// Captured members, by index into the planned parts.
    transfers: Vec<(usize, Option<CompressedPartTransfer>)>,
    /// The destination the candidate is built from: the destination itself,
    /// or the reopen of its serialization when the copy transfers media and
    /// the destination is not an unmodified owned source.
    destination_view: &'a OpcPackage,
}

/// What one candidate build does about its archive and its copied media.
///
/// Carried as one argument because [`build_candidate`] already takes the
/// crate's maximum and `clippy::too_many_arguments` is deny.
#[derive(Clone, Copy)]
struct CandidateBuild<'a> {
    archive: CandidateArchive<'a>,
    media: MediaPolicy,
}

/// An immutable plan to copy one source slide into a different destination
/// presentation.
///
/// The source slide's raw bytes and supported owned dependency closure are
/// retained exactly.  The destination's selected slide supplies the layout
/// boundary; the source layout, master, and theme are never copied or
/// rewritten.  Planning validates a complete candidate before returning.
///
/// The plan is immutable *in value*: every proof it carries is fixed when it
/// is returned, and the one `&mut self` method,
/// [`Self::release_retained_candidate`], gives back memory without changing
/// what the plan proves, publishes or compares equal to.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CrossSlideCopyPlan {
    source: Slide,
    destination: Slide,
    position: usize,
    slide_id: u32,
    presentation_relationship_id: String,
    parts: Box<[super::SlideCopyPart]>,
    source_layout: PackURI,
    destination_layout: PackURI,
    external_relationships: usize,
    planned_bytes: usize,
    source_revision: [u8; 32],
    destination_revision: [u8; 32],
    target_revision: [u8; 32],
    source_physical_revision: [u8; 32],
    destination_physical_revision: [u8; 32],
    target_physical_revision: [u8; 32],
    patch: CrossSlideCopyPatch,
    candidate: RetainedCandidateSlot,
}

/// The serialized candidate archive a plan retains for its own application.
///
/// `archive` is the handle the candidate reopen already holds: planning
/// serializes the candidate into a bounded `Vec`, owned ingress takes that
/// exact allocation as the reopened package's authorized source, and retention
/// takes a second owner of it.  Retention is therefore a decision not to free
/// an allocation, not a second copy of the archive.
///
/// `bound` is the archive bound the bytes were accepted under.  A later build
/// under a different bound is not a reuse.
///
/// `transferred` lists, as ascending indexes into the plan's parts, the copied
/// members whose verified source-compressed bytes the archive frames (change
/// 0742). A later build whose own classification differs is not a reuse
/// either: it serializes its own candidate, which then fails the comparison
/// with the plan exactly as it would without retention.
///
/// This type deliberately does **not** derive `Debug`: a derived one would
/// print the whole archive, and the slot below formats the retained length
/// instead.
#[derive(Clone)]
struct RetainedCandidate {
    archive: Arc<Vec<u8>>,
    bound: usize,
    transferred: Box<[usize]>,
}

/// A plan's retained candidate archive, which is not part of the plan's value.
///
/// The archive is derived state: a plan holding it and the same plan after
/// [`CrossSlideCopyPlan::release_retained_candidate`] prove the same six
/// revisions, carry the same durable patch and publish the same bytes.
/// Equality therefore ignores this slot entirely, which keeps
/// `plan_a == plan_b` the comparison it has always been and makes releasing
/// the archive value-preserving.  `Debug` reports the retained length rather
/// than the bytes.
#[derive(Clone, Default)]
struct RetainedCandidateSlot(Option<RetainedCandidate>);

impl fmt::Debug for RetainedCandidateSlot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("RetainedCandidateSlot")
            .field(
                "retained_bytes",
                &self.0.as_ref().map(|held| held.archive.len()),
            )
            .field("bound", &self.0.as_ref().map(|held| held.bound))
            .field(
                "transferred_members",
                &self.0.as_ref().map(|held| held.transferred.len()),
            )
            .finish()
    }
}

impl PartialEq for RetainedCandidateSlot {
    fn eq(&self, _other: &Self) -> bool {
        true
    }
}

impl Eq for RetainedCandidateSlot {}

/// What one candidate build does about its serialized archive.
///
/// This reaches [`build_candidate`] as a single argument because that function
/// already takes the crate's maximum and `clippy::too_many_arguments` is deny.
#[derive(Clone, Copy)]
enum CandidateArchive<'a> {
    /// Serialize the archive and drop it with the candidate package.
    Build,
    /// Serialize the archive and hand it back for the plan to retain, when the
    /// operation's retained-candidate budget admits its length.
    BuildAndRetain,
    /// Reuse the archive a plan retained under the same bound.
    Reuse(&'a RetainedCandidate),
}

impl CrossSlideCopyPlan {
    /// Bytes of the serialized candidate archive this plan is holding, if any.
    ///
    /// A plan retains the archive it built at planning when its length fits
    /// the intersected [`Limits::max_retained_candidate_bytes`] of the two
    /// snapshots, so that applying the plan can reuse those bytes instead of
    /// serializing the candidate a second time (and capturing or deflating its
    /// copied members again).  `None` means
    /// the plan holds nothing and application rebuilds the archive: either the
    /// candidate was larger than the budget, or the archive was released, or
    /// the candidate reopen did not authorize an exact source.
    ///
    /// The bytes are released by [`Self::release_retained_candidate`] and by
    /// dropping the plan.  Applying a plan does not release them, because a
    /// plan may be applied to more than one destination that proves the same
    /// six revisions.
    #[must_use]
    pub fn retained_candidate_bytes(&self) -> Option<usize> {
        self.candidate.0.as_ref().map(|held| held.archive.len())
    }

    /// Release the retained candidate archive, if this plan holds one.
    ///
    /// The plan keeps its value: it proves the same revisions, carries the
    /// same durable patch, compares equal to the plan it was, and still
    /// applies.  Only the reuse is given up, so a later application rebuilds
    /// and re-serializes the candidate as it does for an unretained plan.
    pub fn release_retained_candidate(&mut self) {
        self.candidate.0 = None;
    }

    /// Source semantic slide captured by the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Slide {
        &self.source
    }

    /// Destination semantic slide whose layout is reused.
    #[must_use]
    pub const fn destination(&self) -> &Slide {
        &self.destination
    }

    /// Checked zero-based insertion position in the destination presentation.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Collision-free destination slide ID reserved by this plan.
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Collision-free destination presentation relationship ID.
    #[must_use]
    pub fn presentation_relationship_id(&self) -> &str {
        &self.presentation_relationship_id
    }

    /// Copied source parts in deterministic source-name order.
    #[must_use]
    pub fn parts(&self) -> &[super::SlideCopyPart] {
        &self.parts
    }

    /// Source layout proved equivalent to the destination layout boundary.
    #[must_use]
    pub const fn source_layout(&self) -> &PackURI {
        &self.source_layout
    }

    /// Existing destination layout reused by the copied slide.
    #[must_use]
    pub const fn destination_layout(&self) -> &PackURI {
        &self.destination_layout
    }

    /// Number of external relationships retained inertly.
    #[must_use]
    pub const fn external_relationship_count(&self) -> usize {
        self.external_relationships
    }

    /// Bounded source-closure and owner bytes inventoried by this plan.
    #[must_use]
    pub const fn planned_bytes(&self) -> usize {
        self.planned_bytes
    }

    /// Complete source-package revision required by the plan.
    #[must_use]
    pub const fn source_revision(&self) -> [u8; 32] {
        self.source_revision
    }

    /// Complete destination-package revision required before publication.
    #[must_use]
    pub const fn destination_revision(&self) -> [u8; 32] {
        self.destination_revision
    }

    /// Complete destination-package revision guaranteed after publication.
    #[must_use]
    pub const fn target_revision(&self) -> [u8; 32] {
        self.target_revision
    }

    /// Exact serialized source-package revision required by the plan.
    #[must_use]
    pub const fn source_physical_revision(&self) -> [u8; 32] {
        self.source_physical_revision
    }

    /// Exact serialized destination-package revision required by the plan.
    #[must_use]
    pub const fn destination_physical_revision(&self) -> [u8; 32] {
        self.destination_physical_revision
    }

    /// Exact serialized destination-package revision guaranteed after publication.
    #[must_use]
    pub const fn target_physical_revision(&self) -> [u8; 32] {
        self.target_physical_revision
    }

    /// Exact durable source- and destination-bound patch.
    #[must_use]
    pub const fn patch(&self) -> &CrossSlideCopyPatch {
        &self.patch
    }

    /// Whether this plan's candidate frames at least one copied image
    /// member from the source member's verified compressed bytes instead of
    /// deflating it again.
    ///
    /// See [`CrossSlideCopyPatch::transfers_source_compressed_media`].
    #[must_use]
    pub fn transfers_source_compressed_media(&self) -> bool {
        self.patch.transfers_source_compressed_media()
    }
}

/// Durable exact cross-presentation copy patch.
///
/// The source revision remains unchanged when the patch is inverted: applying
/// either direction still requires the same immutable source package and the
/// corresponding before-revision of the destination package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CrossSlideCopyPatch {
    source_revision: [u8; 32],
    destination_revision: [u8; 32],
    target_revision: [u8; 32],
    source_physical_revision: [u8; 32],
    destination_physical_revision: [u8; 32],
    target_physical_revision: [u8; 32],
    source_slide: PackURI,
    destination_slide: PackURI,
    destination_layout: PackURI,
    position: usize,
    slide_id: u32,
    presentation_relationship_id: String,
    copied_media: CopiedMedia,
    patch: Patch,
}

impl CrossSlideCopyPatch {
    /// Whether the forward candidate frames at least one copied image member
    /// from the source member's verified compressed bytes instead of
    /// deflating it again.
    ///
    /// Which copied members transfer is decided from the bytes the source
    /// package publishes, never from either package's edit history: a
    /// relationship-free, non-XML `image/*` member is transferred when the
    /// source archive's headers prove its Store or Deflate layout, its
    /// compressed size is at most its decoded size plus stored-block overhead,
    /// and its capture decodes to exactly the planned bytes. A member whose own
    /// bytes fail any of these is deflated again, so the same two packages'
    /// bytes always give the same encoding, which the plan and the durable
    /// patch record.
    ///
    /// Application rebuilds the candidate with the recorded encoding from the
    /// two packages' bytes, after their recorded revisions are checked, so
    /// any source and destination with those revisions — including a
    /// destination restored by this patch's inverse, or one edited back to
    /// the same bytes — publish the same output. A transferring copy into a
    /// destination that is not an unmodified owned source publishes the
    /// candidate reopened from its archive: the destination's save
    /// preferences are carried onto it, but caller-defined `Part`
    /// implementations become built-in parts with the same bytes, where a
    /// recompressing copy keeps them.
    #[must_use]
    pub fn transfers_source_compressed_media(&self) -> bool {
        self.copied_media == CopiedMedia::SourceCompressed
    }

    /// Complete source-package revision required for either direction.
    #[must_use]
    pub const fn source_revision(&self) -> [u8; 32] {
        self.source_revision
    }

    /// Complete destination-package revision required before this direction.
    #[must_use]
    pub const fn destination_revision(&self) -> [u8; 32] {
        self.destination_revision
    }

    /// Complete destination-package revision guaranteed after this direction.
    #[must_use]
    pub const fn target_revision(&self) -> [u8; 32] {
        self.target_revision
    }

    /// Exact serialized source-package revision required for publication.
    #[must_use]
    pub const fn source_physical_revision(&self) -> [u8; 32] {
        self.source_physical_revision
    }

    /// Exact serialized destination-package revision required before publication.
    #[must_use]
    pub const fn destination_physical_revision(&self) -> [u8; 32] {
        self.destination_physical_revision
    }

    /// Exact serialized destination-package revision guaranteed after publication.
    #[must_use]
    pub const fn target_physical_revision(&self) -> [u8; 32] {
        self.target_physical_revision
    }

    /// Number of exact destination resources in the write set.
    #[must_use]
    pub fn resource_count(&self) -> usize {
        self.patch.resource_count()
    }

    /// Physical destination part names in deterministic order.
    pub fn resources(&self) -> impl ExactSizeIterator<Item = &PackURI> {
        self.patch.resources()
    }

    /// Exact inverse direction over the same immutable source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_revision: self.source_revision,
            destination_revision: self.target_revision,
            target_revision: self.destination_revision,
            source_physical_revision: self.source_physical_revision,
            destination_physical_revision: self.target_physical_revision,
            target_physical_revision: self.destination_physical_revision,
            source_slide: self.source_slide.clone(),
            destination_slide: self.destination_slide.clone(),
            destination_layout: self.destination_layout.clone(),
            position: self.position,
            slide_id: self.slide_id,
            presentation_relationship_id: self.presentation_relationship_id.clone(),
            copied_media: self.copied_media,
            patch: self.patch.inverse(),
        }
    }

    /// Serialize this patch into the stable `LPCP0004` binary format.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let payload = self.patch.to_bytes()?;
        let mut output = Vec::new();
        let reserve = PATCH_HEADER_BYTES
            .checked_add(self.source_slide.as_str().len())
            .and_then(|value| value.checked_add(self.destination_slide.as_str().len()))
            .and_then(|value| value.checked_add(self.destination_layout.as_str().len()))
            .and_then(|value| value.checked_add(self.presentation_relationship_id.len()))
            .and_then(|value| value.checked_add(payload.len()))
            .ok_or_else(|| invalid("cross-slide durable patch length overflow"))?;
        let limit = self
            .patch
            .limits()
            .max_patch_bytes()
            .checked_add(PATCH_HEADER_BYTES)
            .ok_or_else(|| invalid("cross-slide durable patch limit overflow"))?;
        if reserve > limit {
            return Err(Error::Limit {
                resource: "cross-slide durable patch bytes",
                limit,
            });
        }
        output
            .try_reserve_exact(reserve)
            .map_err(|source| Error::Allocation {
                resource: "cross-slide durable patch",
                source,
            })?;
        output.extend_from_slice(PATCH_MAGIC);
        output.extend_from_slice(&self.source_revision);
        output.extend_from_slice(&self.destination_revision);
        output.extend_from_slice(&self.target_revision);
        output.extend_from_slice(&self.source_physical_revision);
        output.extend_from_slice(&self.destination_physical_revision);
        output.extend_from_slice(&self.target_physical_revision);
        put_text(&mut output, self.source_slide.as_str(), "source slide")?;
        put_text(
            &mut output,
            self.destination_slide.as_str(),
            "destination slide",
        )?;
        put_text(
            &mut output,
            self.destination_layout.as_str(),
            "destination layout",
        )?;
        put_u64(
            &mut output,
            u64::try_from(self.position)
                .map_err(|_error| invalid("cross-slide insertion position exceeds u64"))?,
        )?;
        put_u32(&mut output, self.slide_id)?;
        put_text(
            &mut output,
            &self.presentation_relationship_id,
            "presentation relationship ID",
        )?;
        output.push(self.copied_media.wire());
        let payload_len = u64::try_from(payload.len())
            .map_err(|_error| invalid("cross-slide durable patch payload exceeds u64"))?;
        put_u64(&mut output, payload_len)?;
        output.extend_from_slice(&payload);
        Ok(output)
    }

    /// Parse a stable durable patch under conservative finite limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DurablePatchRevisionFormat`] for a patch serialized
    /// under the superseded `LPCP0002` or `LPCP0003` format, and an error for
    /// malformed, trailing, or unbounded input.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a stable durable patch under caller-selected finite limits.
    ///
    /// A patch whose header carries a superseded magic is refused here,
    /// before any header field is read. `LPCP0002` embeds three revisions from
    /// the `litchi-pptx-opened-v1` algebra; `LPCP0003` predates the recorded
    /// copied-media encoding that the embedded physical revisions now depend
    /// on, and is refused by name, as change 0655 refused `LPCP0002`, rather
    /// than read implicitly as the recompressed encoding. Re-plan the copy
    /// against the source and destination packages.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DurablePatchRevisionFormat`] for a patch serialized
    /// under the superseded `LPCP0002` or `LPCP0003` format, and an error for
    /// malformed, trailing, or unbounded input.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let limit = limits
            .max_patch_bytes()
            .checked_add(PATCH_HEADER_BYTES)
            .ok_or_else(|| invalid("cross-slide durable patch limit overflow"))?;
        if bytes.len() > limit {
            return Err(Error::Limit {
                resource: "cross-slide durable patch bytes",
                limit,
            });
        }
        // See `SlideRemovalPatch::from_bytes_with_limits`: a superseded magic
        // is refused by name before any header field is read, so a revision
        // from a superseded proof never reaches a caller.
        for superseded in SUPERSEDED_PATCH_FORMATS {
            let magic = superseded.magic();
            if bytes.get(..magic.len()) == Some(magic.as_slice()) {
                return Err(Error::DurablePatchRevisionFormat {
                    found: superseded,
                    expected: PATCH_FORMAT,
                });
            }
        }
        let mut input = WireInput::new(bytes);
        if input.take(PATCH_MAGIC.len())? != PATCH_MAGIC {
            return Err(invalid(
                "cross-slide durable patch has an unsupported version",
            ));
        }
        let source_revision = input.revision()?;
        let destination_revision = input.revision()?;
        let target_revision = input.revision()?;
        let source_physical_revision = input.revision()?;
        let destination_physical_revision = input.revision()?;
        let target_physical_revision = input.revision()?;
        let source_slide = parse_part_name_text(input.text32("source slide")?)?;
        let destination_slide = parse_part_name_text(input.text32("destination slide")?)?;
        let destination_layout = parse_part_name_text(input.text32("destination layout")?)?;
        let position = input.usize64("insertion position")?;
        let slide_id = input.u32()?;
        let presentation_relationship_id = input.text32("presentation relationship ID")?;
        if presentation_relationship_id.is_empty() {
            return Err(invalid(
                "cross-slide durable patch has an empty presentation relationship ID",
            ));
        }
        let copied_media = CopiedMedia::from_wire(input.u8()?)?;
        let payload_len = input.usize64("patch payload")?;
        if payload_len > limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "cross-slide durable patch payload",
                limit: limits.max_patch_bytes(),
            });
        }
        let payload = input.take(payload_len)?;
        if !input.is_empty() {
            return Err(invalid("cross-slide durable patch has trailing bytes"));
        }
        let patch = Patch::from_bytes_with_limits(payload, limits)?;
        validate_patch_descriptor(
            &patch,
            &source_slide,
            &destination_slide,
            &destination_layout,
            position,
            slide_id,
            &presentation_relationship_id,
        )?;
        Ok(Self {
            source_revision,
            destination_revision,
            target_revision,
            source_physical_revision,
            destination_physical_revision,
            target_physical_revision,
            source_slide,
            destination_slide,
            destination_layout,
            position,
            slide_id,
            presentation_relationship_id,
            copied_media,
            patch,
        })
    }
}

impl Snapshot {
    /// Plan a source-checked copy from an immutable source presentation into
    /// this destination presentation.
    ///
    /// `destination_slide` selects the destination layout that will be reused;
    /// `position` selects the insertion point independently.  The source
    /// layout/master/theme inheritance graph must be byte- and
    /// relationship-equivalent to that destination boundary.  Source private
    /// dependencies are copied through the existing strict owned-role closure;
    /// external allowlisted targets remain inert and are never fetched.
    /// Both snapshots must come from source-preserving ingress such as
    /// [`crate::Package::from_vec`], [`crate::Package::open`], or
    /// [`crate::Package::from_reader`]; borrowed graph-only ingress is refused
    /// because it cannot authorize discarded ZIP ordering and extras.
    ///
    /// Each copied relationship-free, non-XML `image/*` member whose source
    /// member's bytes prove a bounded Store or Deflate encoding of it is
    /// published from those verified compressed bytes, in fresh known-size
    /// framing, instead of being deflated again (change 0742). See
    /// [`CrossSlideCopyPatch::transfers_source_compressed_media`] for the rule
    /// and for what it means for application.
    pub fn plan_cross_slide_copy<'s, 'd>(
        &self,
        source: &Snapshot,
        source_slide: impl Into<crate::slide::Key<'s>>,
        destination_slide: impl Into<crate::slide::Key<'d>>,
        position: usize,
    ) -> Result<CrossSlideCopyPlan> {
        let source_slide = resolve_slide(source, source_slide.into())?;
        let destination_slide = resolve_slide(self, destination_slide.into())?;
        plan_cross_slide_copy_for_slides(
            source,
            self,
            source_slide,
            destination_slide,
            position,
            CandidateBuild {
                archive: CandidateArchive::BuildAndRetain,
                media: MediaPolicy::Classify,
            },
        )
    }

    /// Compatibility-oriented alias for [`Self::plan_cross_slide_copy`].
    pub fn plan_cross_presentation_slide_copy<'s, 'd>(
        &self,
        source: &Snapshot,
        source_slide: impl Into<crate::slide::Key<'s>>,
        destination_slide: impl Into<crate::slide::Key<'d>>,
        position: usize,
    ) -> Result<CrossSlideCopyPlan> {
        self.plan_cross_slide_copy(source, source_slide, destination_slide, position)
    }
}

pub(crate) fn apply_plan(
    source: &OpcPackage,
    destination: &mut OpcPackage,
    plan: &CrossSlideCopyPlan,
    source_physical_source_provenance: bool,
    destination_physical_source_provenance: bool,
) -> Result<Snapshot> {
    let limits = plan.patch.patch.limits();
    let source_revision = super::model::package_fingerprint(source)?;
    if source_revision != plan.source_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the complete source package graph changed after cross-slide planning",
        });
    }
    let destination_revision = super::model::package_fingerprint(destination)?;
    if destination_revision != plan.destination_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the complete destination package graph changed after cross-slide planning",
        });
    }
    let source_physical_revision = physical_package_fingerprint(source, limits)?;
    if source_physical_revision != plan.source_physical_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the serialized source package changed after cross-slide planning",
        });
    }
    let destination_physical_revision = physical_package_fingerprint(destination, limits)?;
    if destination_physical_revision != plan.destination_physical_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the serialized destination package changed after cross-slide planning",
        });
    }
    let source_snapshot = super::model::capture_with_revision(
        source,
        limits,
        source_physical_source_provenance,
        source_revision,
    )?;
    remember_physical_revision(&source_snapshot, limits, source_physical_revision);
    let destination_snapshot = super::model::capture_with_revision(
        destination,
        limits,
        destination_physical_source_provenance,
        destination_revision,
    )?;
    remember_physical_revision(&destination_snapshot, limits, destination_physical_revision);
    let (fresh, candidate, normalized) = prepare_cross_slide_copy_for_slides(
        &source_snapshot,
        &destination_snapshot,
        plan.source.clone(),
        plan.destination.clone(),
        plan.position,
        CandidateBuild {
            archive: plan
                .candidate
                .0
                .as_ref()
                .map_or(CandidateArchive::Build, CandidateArchive::Reuse),
            media: MediaPolicy::Recorded(plan.patch.copied_media),
        },
    )?;
    if fresh.source_revision != plan.source_revision
        || fresh.destination_revision != plan.destination_revision
        || fresh.target_revision != plan.target_revision
        || fresh.source_physical_revision != plan.source_physical_revision
        || fresh.destination_physical_revision != plan.destination_physical_revision
        || fresh.target_physical_revision != plan.target_physical_revision
        || fresh.patch != plan.patch
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the durable cross-slide plan does not match a freshly proven candidate",
        });
    }
    let (mut candidate, snapshot, rebuilt) = validate_application_candidate(
        destination,
        candidate,
        &plan.patch.patch,
        plan.target_revision,
        destination_physical_source_provenance,
        !normalized,
    )?;
    if published_archive_revision(&candidate, limits, rebuilt, fresh.target_physical_revision)?
        != plan.target_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the published candidate has an unexpected serialized package revision",
        });
    }
    if normalized {
        adopt_destination_save_options(&mut candidate, destination);
    }
    *destination = candidate;
    Ok(snapshot)
}

pub(crate) fn apply_patch(
    source: &OpcPackage,
    destination: &mut OpcPackage,
    patch: &CrossSlideCopyPatch,
    source_physical_source_provenance: bool,
    destination_physical_source_provenance: bool,
) -> Result<Snapshot> {
    let limits = patch.patch.limits();
    let source_revision = super::model::package_fingerprint(source)?;
    if source_revision != patch.source_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the complete source package graph differs from the cross-slide patch source",
        });
    }
    let destination_revision = super::model::package_fingerprint(destination)?;
    if destination_revision != patch.destination_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the complete destination package graph differs from the cross-slide patch source",
        });
    }
    let source_physical_revision = physical_package_fingerprint(source, limits)?;
    if source_physical_revision != patch.source_physical_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the serialized source package differs from the cross-slide patch source",
        });
    }
    let destination_physical_revision = physical_package_fingerprint(destination, limits)?;
    if destination_physical_revision != patch.destination_physical_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the serialized destination package differs from the cross-slide patch source",
        });
    }
    let source_snapshot = super::model::capture_with_revision(
        source,
        limits,
        source_physical_source_provenance,
        source_revision,
    )?;
    remember_physical_revision(&source_snapshot, limits, source_physical_revision);
    let destination_snapshot = super::model::capture_with_revision(
        destination,
        limits,
        destination_physical_source_provenance,
        destination_revision,
    )?;
    remember_physical_revision(&destination_snapshot, limits, destination_physical_revision);
    let source_slide = find_slide_by_part(&source_snapshot, &patch.source_slide)?;
    let destination_slide = find_slide_by_part(&destination_snapshot, &patch.destination_slide)?;
    let forward_candidate = prepare_cross_slide_copy_for_slides(
        &source_snapshot,
        &destination_snapshot,
        source_slide.clone(),
        destination_slide.clone(),
        patch.position,
        CandidateBuild {
            archive: CandidateArchive::Build,
            media: MediaPolicy::Recorded(patch.copied_media),
        },
    )
    .ok()
    .and_then(|(fresh, candidate, normalized)| {
        (fresh.patch == *patch
            && fresh.target_revision == patch.target_revision
            && fresh.target_physical_revision == patch.target_physical_revision
            && fresh.slide_id == patch.slide_id
            && fresh.presentation_relationship_id == patch.presentation_relationship_id)
            .then_some((candidate, true, normalized))
    });
    let candidate = if let Some(candidate) = forward_candidate {
        Some(candidate)
    } else {
        // An inverse patch is validated by restoring a detached candidate to
        // its forward source revision, freshly replanning the forward copy,
        // and comparing the exact inverse. The real destination is untouched
        // until this proof succeeds.
        let mut restored = destination.clone();
        if super::patch::apply_exact_revision(
            &mut restored,
            &patch.patch,
            patch.target_revision,
            destination_physical_source_provenance,
        )
        .is_err()
        {
            None
        } else if physical_package_fingerprint(&restored, limits).ok()
            != Some(patch.target_physical_revision)
        {
            None
        } else {
            let restored_snapshot =
                super::model::capture(&restored, limits, destination_physical_source_provenance);
            let inverse_matches = restored_snapshot
                .ok()
                .inspect(|base| {
                    remember_physical_revision(base, limits, patch.target_physical_revision);
                })
                .and_then(|base| {
                    let restored_destination =
                        find_slide_by_part(&base, &patch.destination_slide).ok()?;
                    let forward = plan_cross_slide_copy_for_slides(
                        &source_snapshot,
                        &base,
                        source_slide,
                        restored_destination,
                        patch.position,
                        CandidateBuild {
                            archive: CandidateArchive::Build,
                            media: MediaPolicy::Recorded(patch.copied_media),
                        },
                    )
                    .ok()?;
                    Some(forward.patch.inverse() == *patch)
                })
                .unwrap_or(false);
            inverse_matches.then_some((restored, false, false))
        }
    };
    let (candidate, reopened, normalized) = candidate.ok_or(Error::UnsafeEdit {
        operation: "apply_cross_slide_copy_patch",
        reason: "the durable cross-slide patch does not match a freshly proven candidate",
    })?;
    let (mut candidate, snapshot, rebuilt) = validate_application_candidate(
        destination,
        candidate,
        &patch.patch,
        patch.target_revision,
        destination_physical_source_provenance,
        reopened && !normalized,
    )?;
    if published_archive_revision(&candidate, limits, rebuilt, patch.target_physical_revision)?
        != patch.target_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the published candidate has an unexpected serialized package revision",
        });
    }
    if normalized {
        adopt_destination_save_options(&mut candidate, destination);
    }
    *destination = candidate;
    Ok(snapshot)
}

/// Validate the candidate application publishes, or rebuild it from the live
/// destination.
///
/// `may_rebuild` is false when the candidate must be published as built: the
/// inverse route's restored clone, and a transferring copy whose candidate
/// was built from the destination's reopened serialization, which a
/// clone-and-apply rebuild could not reproduce because the targeted writer
/// deflates the patch's decoded resources again.
fn validate_application_candidate(
    destination: &OpcPackage,
    candidate: OpcPackage,
    patch: &Patch,
    target_revision: [u8; 32],
    physical_source_provenance: bool,
    may_rebuild: bool,
) -> Result<(OpcPackage, Snapshot, bool)> {
    // Reopening preserves the observable state of untouched owned ingress.
    // Dirty packages can carry caller-defined parts or save preferences that
    // are absent from the archive, so retain their clone-and-apply behavior
    // whenever the candidate is a recompressed one.
    if may_rebuild && !destination.is_unmodified_owned_source() {
        drop(candidate);
        let mut candidate = destination.clone();
        let snapshot = super::patch::apply_exact_revision(
            &mut candidate,
            patch,
            target_revision,
            physical_source_provenance,
        )?;
        return Ok((candidate, snapshot, true));
    }
    let snapshot = super::patch::validate_candidate(
        destination,
        &candidate,
        patch,
        target_revision,
        physical_source_provenance,
    )?;
    Ok((candidate, snapshot, false))
}

/// Carry a modified destination's save preferences onto a transferring
/// copy's published candidate, which is a reopen of the candidate archive.
///
/// Save preferences are not part of any archive, so the reopen cannot carry
/// them itself; they do not change what the package serializes to, so the
/// proven physical revision still holds. Default preferences are not set,
/// which keeps an unmodified destination's candidate an unmodified owned
/// source. Caller-defined `Part` implementations are not carried: the
/// candidate holds built-in parts with the same bytes.
fn adopt_destination_save_options(candidate: &mut OpcPackage, destination: &OpcPackage) {
    // Destructured, so a new save preference cannot be missed here.
    let litchi_opc::SaveOptions { fonts } = destination.save_options();
    if *fonts != litchi_opc::FontEmbedding::None {
        candidate.set_save_options(destination.save_options().clone());
    }
}

/// Serialized-archive revision of the package application is about to publish.
///
/// When the candidate reaching publication is the one whose archive revision
/// was proven a moment earlier — the prepared candidate in `apply_plan` and in
/// the forward route of `apply_patch`, or the restored candidate in the inverse
/// route — the value is already known and recomputing it can only reproduce it.
/// A candidate `validate_application_candidate` rebuilt from the destination is
/// a different package and is hashed.
fn published_archive_revision(
    candidate: &OpcPackage,
    limits: Limits,
    rebuilt: bool,
    proven: [u8; 32],
) -> Result<[u8; 32]> {
    if rebuilt {
        return physical_package_fingerprint(candidate, limits);
    }
    debug_assert!(
        physical_package_fingerprint(candidate, limits).is_ok_and(|fresh| fresh == proven),
        "cross-slide published a candidate with an unproven serialized package revision"
    );
    Ok(proven)
}

fn plan_cross_slide_copy_for_slides(
    source: &Snapshot,
    destination: &Snapshot,
    source_slide: Slide,
    destination_slide: Slide,
    position: usize,
    build: CandidateBuild<'_>,
) -> Result<CrossSlideCopyPlan> {
    prepare_cross_slide_copy_for_slides(
        source,
        destination,
        source_slide,
        destination_slide,
        position,
        build,
    )
    .map(|(plan, _candidate, _normalized)| plan)
}

// Keep the reopened candidate only within an application call. Public plans
// retain their existing descriptor and durable patch representation.
fn prepare_cross_slide_copy_for_slides(
    source: &Snapshot,
    destination: &Snapshot,
    source_slide: Slide,
    destination_slide: Slide,
    position: usize,
    build: CandidateBuild<'_>,
) -> Result<(CrossSlideCopyPlan, OpcPackage, bool)> {
    if position > destination.slides.len() {
        return Err(Error::SlideIndexOutOfBounds {
            index: position,
            len: destination.slides.len().saturating_add(1),
        });
    }
    if !source.physical_source_provenance || !destination.physical_source_provenance {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            "cross-slide physical authorization requires source-preserving package ingress (use Package::from_vec, open, or from_reader)",
        );
    }
    if has_signature_infrastructure(source.package.as_ref())
        || has_signature_infrastructure(destination.package.as_ref())
    {
        return refusal(
            SlideCopyRefusal::SignedPackage,
            "digital-signature infrastructure requires an explicit signature policy",
        );
    }
    reject_unknown_non_part_members(source.package.as_ref(), "cross-slide source")?;
    reject_unknown_non_part_members(destination.package.as_ref(), "cross-slide destination")?;
    if has_macro_infrastructure(source.package.as_ref())
        || has_macro_infrastructure(destination.package.as_ref())
    {
        return refusal(
            SlideCopyRefusal::UnsupportedRelationship,
            "macro/VBA infrastructure is outside cross-presentation slide copying",
        );
    }
    let source_presentation = source.package.get_part(&source.presentation_name)?;
    let destination_presentation = destination
        .package
        .get_part(&destination.presentation_name)?;
    let source_dialect = prove_package_dialect(source.package.as_ref(), source_presentation)?;
    let destination_dialect =
        prove_package_dialect(destination.package.as_ref(), destination_presentation)?;
    if source_dialect != destination_dialect {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "source and destination PresentationML packages use different strict/transitional dialects",
        );
    }
    reject_mce(source_presentation.blob(), "source presentation owner")?;
    reject_mce(
        destination_presentation.blob(),
        "destination presentation owner",
    )?;
    reject_protected(source_presentation.blob(), "source presentation")?;
    reject_protected(destination_presentation.blob(), "destination presentation")?;
    if source.slides.is_empty() || destination.slides.is_empty() {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "cross-slide copy requires a source and destination presentation with slides",
        );
    }
    reject_slide_name_collisions(&source.slides, &destination.slides, &source_slide)?;
    if destination.slides.len() == crate::parts::MAX_SLIDES {
        return Err(Error::Limit {
            resource: "cross-slide copy presentation slides",
            limit: crate::parts::MAX_SLIDES,
        });
    }
    validate_slide_ids(&source.slides)?;
    validate_slide_ids(&destination.slides)?;
    let limits = intersect_limits(source.limits, destination.limits)?;
    if source.package.part_count() > limits.max_parts() {
        return Err(Error::Limit {
            resource: "cross-slide copy source package parts",
            limit: limits.max_parts(),
        });
    }
    if destination.package.part_count() > limits.max_parts() {
        return Err(Error::Limit {
            resource: "cross-slide copy destination package parts",
            limit: limits.max_parts(),
        });
    }
    crate::master_layout::validate_master_layout_graph(source.package.as_ref())?;
    crate::master_layout::validate_master_layout_graph(destination.package.as_ref())?;

    let source_part = source.package.get_part(&source_slide.part_name)?;
    crate::parts::validate_content_type(source_part, ct::PML_SLIDE)?;
    validate_slide_surface(source_part.blob())?;
    let destination_part = destination.package.get_part(&destination_slide.part_name)?;
    crate::parts::validate_content_type(destination_part, ct::PML_SLIDE)?;
    let source_layout = selected_layout(source.package.as_ref(), source_part)?;
    let destination_layout = selected_layout(destination.package.as_ref(), destination_part)?;
    validate_registered_layout(source.package.as_ref(), source_presentation, &source_layout)?;
    validate_registered_layout(
        destination.package.as_ref(),
        destination_presentation,
        &destination_layout,
    )?;
    prove_layout_inheritance(
        source.package.as_ref(),
        &source_layout,
        destination.package.as_ref(),
        &destination_layout,
        limits,
    )?;

    let (owned, edges, _reused, external_relationships, planned_bytes) =
        collect_owned_closure(source.package.as_ref(), &source_slide.part_name, limits)?;
    super::copy_plan::reject_cycles(&owned, &edges)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(owned.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide source closure names",
            source,
        })?;
    names.extend(owned);
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(names.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide planned parts",
            source,
        })?;
    let mut reserved = HashSet::new();
    reserved
        .try_reserve(names.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide target identities",
            source,
        })?;
    for name in names {
        let part = source.package.get_part(&name)?;
        let target = available_copy_name_avoiding(
            destination.package.as_ref(),
            &name,
            limits.max_parts(),
            &reserved,
        )?;
        reserved.insert(target.clone());
        parts.push(super::SlideCopyPart {
            source: name,
            target,
            content_type: copy_string(part.content_type(), "cross-slide content types")?,
            bytes: part.blob().len(),
            relationships: part.rels().len(),
        });
    }
    let resulting_parts = destination
        .package
        .part_count()
        .checked_add(parts.len())
        .ok_or_else(|| invalid("cross-slide resulting part count overflow"))?;
    if resulting_parts > limits.max_parts() {
        return Err(Error::Limit {
            resource: "cross-slide resulting package parts",
            limit: limits.max_parts(),
        });
    }
    if planned_bytes > limits.max_patch_bytes() {
        return Err(Error::Limit {
            resource: "cross-slide planned closure bytes",
            limit: limits.max_patch_bytes(),
        });
    }
    let slide_id = next_slide_id(&destination.slides)?;
    let presentation_relationship_id = next_relationship_id(destination_presentation.rels())?;
    let media = classify_media(source, &parts, limits, build.media)?;
    preflight_parts(destination, &parts, planned_bytes, media.capture_bytes)?;
    let reuse = match build.archive {
        CandidateArchive::Reuse(held) => Some(held),
        CandidateArchive::Build | CandidateArchive::BuildAndRetain => None,
    };
    let media = media.capture(&parts, reuse, limits.max_patch_bytes())?;
    // A transferring copy is built from the bytes the destination publishes,
    // so its candidate, and the package application publishes, never depend
    // on the destination's edit history. A recompressing copy keeps building
    // from the destination itself, whose clone-and-apply publication retains
    // caller-defined parts and save preferences.
    let destination_view = if media.members.is_empty() {
        Cow::Borrowed(destination.package.as_ref())
    } else {
        owned_view(destination, limits)?
    };
    let normalized = matches!(destination_view, Cow::Owned(_));
    let source_physical_revision = snapshot_physical_revision(source, limits)?;
    let destination_physical_revision = snapshot_physical_revision(destination, limits)?;
    let BuiltCandidate {
        candidate,
        revision: candidate_revision,
        archive_revision: candidate_archive_revision,
        retained: retained_candidate,
        copied_media,
    } = build_candidate(
        source,
        destination,
        &source_slide,
        &destination_slide,
        position,
        slide_id,
        &presentation_relationship_id,
        &source_layout,
        &destination_layout,
        &parts,
        limits,
        CandidateInputs {
            archive: build.archive,
            transfers: media.members,
            destination_view: &destination_view,
        },
    )?;
    let patch = Patch::capture(
        destination.package.as_ref(),
        &candidate,
        destination.presentation_name.clone(),
        limits,
    )?;
    // `build_candidate` already captured this exact package, and a capture's
    // revision is `package_fingerprint` of the package it captured.
    let target_revision = candidate_revision;
    let target_physical_revision =
        candidate_physical_revision(&candidate, limits, candidate_archive_revision)?;
    let cross_patch = CrossSlideCopyPatch {
        source_revision: source.revision,
        destination_revision: destination.revision,
        target_revision,
        source_physical_revision,
        destination_physical_revision,
        target_physical_revision,
        source_slide: source_slide.part_name.clone(),
        destination_slide: destination_slide.part_name.clone(),
        destination_layout: destination_layout.clone(),
        position,
        slide_id,
        presentation_relationship_id: presentation_relationship_id.clone(),
        copied_media,
        patch,
    };
    validate_patch_descriptor(
        &cross_patch.patch,
        &cross_patch.source_slide,
        &cross_patch.destination_slide,
        &cross_patch.destination_layout,
        cross_patch.position,
        cross_patch.slide_id,
        &cross_patch.presentation_relationship_id,
    )?;
    Ok((
        CrossSlideCopyPlan {
            source: source_slide,
            destination: destination_slide,
            position,
            slide_id,
            presentation_relationship_id,
            parts: parts.into_boxed_slice(),
            source_layout,
            destination_layout,
            external_relationships,
            planned_bytes,
            source_revision: source.revision,
            destination_revision: destination.revision,
            target_revision,
            source_physical_revision,
            destination_physical_revision,
            target_physical_revision,
            patch: cross_patch,
            candidate: retained_candidate,
        },
        candidate,
        normalized,
    ))
}

/// One built, reopened and captured candidate.
struct BuiltCandidate {
    /// The candidate reopened from its serialized archive.
    candidate: OpcPackage,
    /// Complete-package revision of `candidate`.
    revision: [u8; 32],
    /// Physical revision of the archive when the reopen republishes it
    /// verbatim.
    archive_revision: Option<[u8; 32]>,
    retained: RetainedCandidateSlot,
    copied_media: CopiedMedia,
}

/// Whether one copied member may carry its source member's compressed bytes.
///
/// The format-level half of the rule: a relationship-free `image/*` part that
/// is not XML by name or content type (an SVG is `image/svg+xml` and is XML).
/// The package-level half — owned source archive, source member present,
/// payload still the opened allocation, content type unchanged, no signature
/// infrastructure, a provable layout and a bounded compressed size — is
/// [`OpcPackage::compressed_transfer_size`], which reads only provenance and
/// central-directory metadata. This half reads a part planning has already
/// decoded.
fn is_transferable_media(part: &dyn Part) -> bool {
    let content_type = part.content_type();
    content_type
        .get(..6)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("image/"))
        && !is_xml_part(part.partname(), content_type)
        && part.rels().is_empty()
}

/// Apply the format rule and the source's header-only eligibility to every
/// planned part.
///
/// A recorded recompressed encoding transfers nothing. Otherwise every part
/// that passes the format rule is asked about through the package whose
/// retained archive is exactly what the source publishes (see
/// [`owned_view`]), so the answer depends on the source's bytes alone. No
/// member is captured here, so every limit can be checked first.
fn classify_media<'s>(
    source: &'s Snapshot,
    parts: &[super::SlideCopyPart],
    limits: Limits,
    policy: MediaPolicy,
) -> Result<MediaTransfers<'s>> {
    let transfer = match policy {
        MediaPolicy::Classify => true,
        MediaPolicy::Recorded(recorded) => recorded == CopiedMedia::SourceCompressed,
    };
    let mut candidates = Vec::new();
    if transfer {
        candidates
            .try_reserve_exact(parts.len())
            .map_err(|source| Error::Allocation {
                resource: "cross-slide transferred media",
                source,
            })?;
        for (index, planned) in parts.iter().enumerate() {
            if is_transferable_media(source.package.get_part(&planned.source)?) {
                candidates.push(index);
            }
        }
    }
    if candidates.is_empty() {
        return Ok(MediaTransfers {
            source_view: Cow::Borrowed(source.package.as_ref()),
            members: Vec::new(),
            capture_bytes: 0,
        });
    }
    let source_view = owned_view(source, limits)?;
    let mut members = Vec::new();
    members
        .try_reserve_exact(candidates.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide transferred media",
            source,
        })?;
    let mut capture_bytes = 0usize;
    for index in candidates {
        let Some(size) = source_view.compressed_transfer_size(&parts[index].source)? else {
            continue;
        };
        capture_bytes = usize::try_from(size)
            .ok()
            .and_then(|size| capture_bytes.checked_add(size))
            .ok_or_else(|| invalid("cross-slide transferred media byte count overflow"))?;
        members.push((index, None));
    }
    Ok(MediaTransfers {
        source_view,
        members,
        capture_bytes,
    })
}

impl MediaTransfers<'_> {
    /// Capture every eligible member, keeping those whose capture verifies.
    ///
    /// A member whose own bytes disprove its capture — for example a Deflate
    /// stream followed by bytes it does not consume, which the ordinary reader
    /// tolerates — is recompressed instead: the same bytes always give the
    /// same answer, so planning and every later proof agree. Limits,
    /// allocation, I/O and cancellation stay typed errors.
    ///
    /// A release build that reuses a retained archive built from exactly
    /// these members skips the captures: the archive framed them when it was
    /// planned, and both packages are proven unchanged since.
    fn capture(
        self,
        parts: &[super::SlideCopyPart],
        reuse: Option<&RetainedCandidate>,
        archive_limit: usize,
    ) -> Result<Self> {
        let Self {
            source_view,
            members,
            capture_bytes,
        } = self;
        let reusable = reuse.is_some_and(|held| {
            held.bound == archive_limit
                && held.archive.len() <= archive_limit
                && held
                    .transferred
                    .iter()
                    .copied()
                    .eq(members.iter().map(|(index, _capture)| *index))
        });
        if reusable && cfg!(not(debug_assertions)) {
            return Ok(Self {
                source_view,
                members,
                capture_bytes,
            });
        }
        let mut captured = Vec::new();
        captured
            .try_reserve_exact(members.len())
            .map_err(|source| Error::Allocation {
                resource: "cross-slide transferred media",
                source,
            })?;
        for (index, _capture) in members {
            let planned = &parts[index];
            let Some(capture) = source_view.authorize_compressed_transfer(&planned.source)? else {
                continue;
            };
            if capture.content_type() != planned.content_type {
                return Err(invalid(
                    "cross-slide compressed transfer changed the copied content type",
                ));
            }
            captured.push((index, Some(capture)));
        }
        Ok(Self {
            source_view,
            members: captured,
            capture_bytes,
        })
    }
}

/// The package that holds, as its owned source archive, exactly the bytes
/// `snapshot` publishes.
///
/// An unmodified owned source already does and is returned as it is. Any
/// other package — one edited since it was opened, or the restored clone an
/// inverse application publishes — is serialized under the operation's
/// archive bound and reopened as an owned source, so every later decision
/// about its members is a function of the bytes it publishes, never of its
/// edit history or of allocation identity. The serialization is the one its
/// physical revision hashes, bounded and charged by the same writer against
/// `max_patch_bytes`, so no package whose revision could be taken is refused
/// for size here; the snapshot's recorded revision is checked against it, or
/// seeded from it. The reopen re-admits the bytes under the read limits the
/// package's own archive was admitted under, so the view's captures are held
/// to the policy an unmodified package's captures are held to.
fn owned_view(snapshot: &Snapshot, limits: Limits) -> Result<Cow<'_, OpcPackage>> {
    let package = snapshot.package.as_ref();
    if package.is_unmodified_owned_source() {
        return Ok(Cow::Borrowed(package));
    }
    reject_unknown_non_part_members(package, "cross-slide physical authorization")?;
    let bound = limits.max_patch_bytes();
    let (bytes, digest) = bounded_package_bytes(package, bound)?;
    let revision = seal_physical_revision(digest, bytes.len())?;
    match snapshot.physical_revision.get() {
        Some(&(cached_bound, cached)) if cached_bound == bound => {
            if cached != revision {
                return Err(invalid(
                    "cross-slide serialized a package differently from its physical revision",
                ));
            }
        },
        _ => {
            let _first = snapshot.physical_revision.set((bound, revision));
        },
    }
    let view =
        OpcPackage::from_vec_with_limits(bytes, package.source_read_limits().unwrap_or_default())?;
    if !view.is_unmodified_owned_source() {
        return Err(invalid(
            "cross-slide could not reopen a package's serialization as an owned source",
        ));
    }
    Ok(Cow::Owned(view))
}

fn build_candidate(
    source: &Snapshot,
    destination: &Snapshot,
    source_slide: &Slide,
    destination_slide: &Slide,
    position: usize,
    slide_id: u32,
    presentation_relationship_id: &str,
    source_layout: &PackURI,
    destination_layout: &PackURI,
    parts: &[super::SlideCopyPart],
    limits: Limits,
    inputs: CandidateInputs<'_>,
) -> Result<BuiltCandidate> {
    let CandidateInputs {
        archive,
        transfers,
        destination_view,
    } = inputs;
    let archive_limit = limits.max_patch_bytes();
    let mut transferred = Vec::new();
    transferred
        .try_reserve_exact(transfers.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide transferred media",
            source,
        })?;
    transferred.extend(transfers.iter().map(|(index, _capture)| *index));
    let reuses_archive = matches!(
        archive,
        CandidateArchive::Reuse(held)
            if held.bound == archive_limit
                && held.archive.len() <= archive_limit
                && *held.transferred == *transferred
    );
    // A member is uncaptured only when a retained archive framing it is
    // reused, so the graph built below is never serialized with it.
    if !reuses_archive && transfers.iter().any(|(_index, capture)| capture.is_none()) {
        return Err(invalid(
            "cross-slide candidate would serialize an uncaptured transferred member",
        ));
    }
    let mut mapping = HashMap::new();
    mapping
        .try_reserve(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide part-name mapping",
            source,
        })?;
    for part in parts {
        if mapping
            .insert(part.source.clone(), part.target.clone())
            .is_some()
        {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "the cross-slide closure repeats a source part",
            );
        }
    }
    let copied_slide = mapping
        .get(&source_slide.part_name)
        .cloned()
        .ok_or_else(|| invalid("cross-slide candidate omitted the selected source slide"))?;
    let mut candidate = destination_view.clone();
    let mut captures = transfers.into_iter().peekable();
    for (index, planned) in parts.iter().enumerate() {
        let original = source.package.get_part(&planned.source)?;
        let capture = captures
            .next_if(|(transferred, _capture)| *transferred == index)
            .and_then(|(_index, capture)| capture);
        let mut copied = match capture {
            Some(capture) => BlobPart::with_compressed_transfer(planned.target.clone(), capture),
            // A recompressed member, or a transferred one whose retained
            // archive is reused: its decoded bytes are the capture's.
            None => BlobPart::new_shared(
                planned.target.clone(),
                planned.content_type.clone(),
                original.blob_arc(),
            ),
        };
        for relationship in original.rels().iter() {
            let (target, mode) = if relationship.is_external() {
                (relationship.target_ref().to_owned(), TargetMode::External)
            } else {
                let source_target = relationship.target_partname()?;
                let target_part = if planned.source == source_slide.part_name
                    && source_target == *source_layout
                {
                    destination_layout
                } else {
                    mapping
                        .get(&source_target)
                        .ok_or_else(|| Error::SlideCopyPlan {
                            kind: SlideCopyRefusal::AmbiguousTopology,
                            detail: "a copied internal relationship escaped the owned closure"
                                .to_owned(),
                        })?
                };
                (
                    target_part.relative_ref(planned.target.base_uri()),
                    TargetMode::Internal,
                )
            };
            copied.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                target,
                relationship.r_id().to_owned(),
                mode,
            )?;
        }
        candidate.try_add_part(Box::new(copied))?;
    }
    let presentation = destination_view.get_part(&destination.presentation_name)?;
    let destination_relationship = presentation
        .rels()
        .get(&destination_slide.relationship_id)
        .ok_or_else(|| invalid("cross-slide destination slide relationship disappeared"))?;
    let xml = super::xml::insert_slide(
        presentation.blob(),
        &destination.slides,
        position,
        slide_id,
        presentation_relationship_id,
    )?;
    {
        let staged = candidate.get_part_mut(&destination.presentation_name)?;
        staged.rels_mut().try_add_relationship(
            destination_relationship.reltype().to_owned(),
            copied_slide.relative_ref(destination.presentation_name.base_uri()),
            presentation_relationship_id.to_owned(),
            TargetMode::Internal,
        )?;
        staged.set_blob(xml);
    }
    // A retained candidate archive is the serialization of this exact graph.
    // Both input packages were proved unchanged before this call -- semantic
    // graph and serialized archive alike -- and `to_stream` is a deterministic
    // function of the package, so a fresh serialization reproduces the
    // retained bytes. The `debug_assert` re-derives exactly that on every
    // reuse in debug and test builds; a release build rests on the argument
    // plus the recomputed archive and graph revisions taken below from the
    // bytes it is about to publish.
    let (serialized, archive_digest) = match archive {
        CandidateArchive::Reuse(held) if reuses_archive => {
            debug_assert!(
                bounded_package_bytes(&candidate, archive_limit)
                    .is_ok_and(|(fresh, _)| fresh == *held.archive),
                "cross-slide reused a retained candidate archive a fresh serialization does not reproduce"
            );
            // The copy is fallible, exactly as `BoundedVecWriter::into_bytes`
            // is: an infallible `Vec::clone` would abort the process where the
            // route it replaces returns `Error::Allocation`.
            let mut serialized = Vec::new();
            serialized
                .try_reserve_exact(held.archive.len())
                .map_err(|source| Error::Allocation {
                    resource: "cross-slide candidate archive",
                    source,
                })?;
            serialized.extend_from_slice(&held.archive);
            // The digest is recomputed over the retained bytes rather than
            // carried beside them, so the sealed physical revision stays a
            // hash of the bytes this call publishes. That hash is the one
            // `bounded_package_bytes` takes anyway, so what reuse removes is
            // exactly the serialization with its deflate and, in release
            // builds, the copied media captures.
            let mut digest = Sha256::new();
            digest.update(&serialized);
            (serialized, digest.finalize().into())
        },
        _ => bounded_package_bytes(&candidate, archive_limit)?,
    };
    let serialized_bytes = serialized.len();
    // Clean owned ingress proves the destination has built-in parts. Keep
    // the existing path for caller-defined parts and revoked authorization.
    let reopened = if destination_view.is_unmodified_owned_source() {
        OpcPackage::from_vec_reusing_payloads(
            serialized,
            litchi_opc::ReadLimits::default(),
            &candidate,
        )?
    } else {
        OpcPackage::from_vec(serialized)?
    };
    // Owned ingress retains the archive it was opened from and republishes it
    // verbatim, so the bytes just hashed are exactly what a serializing hash
    // sink would read back. A reopen that revoked that authorization is left
    // to the ordinary recomputation.
    let archive_revision = reopened
        .is_unmodified_owned_source()
        .then(|| seal_physical_revision(archive_digest, serialized_bytes))
        .transpose()?;
    let copied_media = if transferred.is_empty() {
        CopiedMedia::Recompressed
    } else {
        CopiedMedia::SourceCompressed
    };
    // Owned ingress took the very allocation `serialized` occupied and keeps
    // it behind a shared handle, so retention is a second owner of those
    // bytes rather than a second copy of them. A candidate above the
    // operation's retained-candidate budget is simply not retained: rebuilding
    // is always available, so the budget is a ceiling on what may be held and
    // never a reason to refuse a plan.
    let retained_candidate = RetainedCandidateSlot(
        (matches!(archive, CandidateArchive::BuildAndRetain)
            && serialized_bytes <= limits.max_retained_candidate_bytes())
        .then(|| reopened.exact_source_shared())
        .flatten()
        .map(|shared| RetainedCandidate {
            archive: shared,
            bound: archive_limit,
            transferred: transferred.into_boxed_slice(),
        }),
    );
    let captured = super::model::capture(
        &reopened,
        destination.limits,
        destination.physical_source_provenance,
    )?;
    let published = captured
        .slides
        .get(position)
        .ok_or_else(|| invalid("cross-slide candidate lost its insertion position"))?;
    if published.id != slide_id || published.part_name != copied_slide {
        return Err(invalid(
            "cross-slide candidate did not publish the reserved slide identity",
        ));
    }
    let copied_part = reopened.get_part(&copied_slide)?;
    let layout = selected_layout(&reopened, copied_part)?;
    if layout != *destination_layout {
        return Err(invalid(
            "cross-slide candidate did not retarget the copied slide layout",
        ));
    }
    let revision = captured.revision;
    Ok(BuiltCandidate {
        candidate: reopened,
        revision,
        archive_revision,
        retained: retained_candidate,
        copied_media,
    })
}

/// Bytes a candidate build is charged against the destination's
/// `max_patch_bytes` before anything is built: twice the planned closure and
/// presentation owner, the new names and content types, a constant, and the
/// compressed bytes the transferred members' captures hold.
fn candidate_estimate(
    destination: &Snapshot,
    parts: &[super::SlideCopyPart],
    planned_bytes: usize,
    capture_bytes: usize,
) -> Result<usize> {
    let owner = destination
        .package
        .get_part(&destination.presentation_name)?;
    planned_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(owner.blob().len().checked_mul(2)?))
        .and_then(|value| {
            parts.iter().try_fold(value, |total, part| {
                total
                    .checked_add(part.target.as_str().len())
                    .and_then(|next| next.checked_add(part.content_type.len()))
            })
        })
        .and_then(|value| value.checked_add(128))
        .and_then(|value| value.checked_add(capture_bytes))
        .ok_or_else(|| invalid("cross-slide candidate byte count overflow"))
}

fn preflight_parts(
    destination: &Snapshot,
    parts: &[super::SlideCopyPart],
    planned_bytes: usize,
    capture_bytes: usize,
) -> Result<()> {
    let estimate = candidate_estimate(destination, parts, planned_bytes, capture_bytes)?;
    if estimate > destination.limits.max_patch_bytes() {
        return Err(Error::Limit {
            resource: "cross-slide candidate patch bytes",
            limit: destination.limits.max_patch_bytes(),
        });
    }
    let mut targets = HashSet::new();
    targets
        .try_reserve(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide target identities",
            source,
        })?;
    for part in parts {
        destination.package.validate_new_part_name(&part.target)?;
        if destination
            .package
            .non_part_members()
            .iter()
            .any(|member| member.name().eq_ignore_ascii_case(part.target.membername()))
        {
            return refusal(
                SlideCopyRefusal::UnknownPhysicalMember,
                "a copied destination Part name collides with an unknown raw ZIP member",
            );
        }
        if !targets.insert(part.target.clone()) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "two copied source resources selected the same destination part name",
            );
        }
    }
    Ok(())
}

fn selected_layout(package: &OpcPackage, slide: &dyn Part) -> Result<PackURI> {
    let mut selected = None;
    for relationship in slide.rels().iter() {
        if !crate::parts::is_relationship_type(
            relationship.reltype(),
            rt::SLIDE_LAYOUT,
            "slideLayout",
        ) {
            continue;
        }
        if relationship.is_external()
            || relationship.target_mode() != TargetMode::Internal
            || relationship.target_query().is_some()
            || relationship.target_fragment().is_some()
        {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "a slide layout relationship is external or has a query/fragment",
            );
        }
        if selected.is_some() {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "a slide has more than one layout relationship",
            );
        }
        let target = relationship.target_partname()?;
        crate::parts::validate_content_type(package.get_part(&target)?, ct::PML_SLIDE_LAYOUT)?;
        selected = Some(target);
    }
    selected.ok_or_else(|| Error::SlideCopyPlan {
        kind: SlideCopyRefusal::AmbiguousTopology,
        detail: "a slide has no reusable layout relationship".to_owned(),
    })
}

#[derive(Clone, Copy)]
enum InheritanceSurface {
    Layout,
    Master,
    Theme,
}

fn reject_inheritance_edges(
    part: &dyn Part,
    surface: InheritanceSurface,
    label: &'static str,
) -> Result<()> {
    for relationship in part.rels().iter() {
        if relationship.is_external()
            || relationship.target_query().is_some()
            || relationship.target_fragment().is_some()
            || relationship.target_mode() != TargetMode::Internal
        {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{label} contains a non-exact inheritance target"),
            );
        }
        let allowed = match surface {
            InheritanceSurface::Layout => crate::parts::is_relationship_type(
                relationship.reltype(),
                rt::SLIDE_MASTER,
                "slideMaster",
            ),
            InheritanceSurface::Master => {
                crate::parts::is_relationship_type(
                    relationship.reltype(),
                    rt::SLIDE_LAYOUT,
                    "slideLayout",
                ) || crate::parts::is_relationship_type(relationship.reltype(), rt::THEME, "theme")
            },
            InheritanceSurface::Theme => false,
        };
        if !allowed {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                format!("{label} contains an unsupported inheritance relationship"),
            );
        }
    }
    Ok(())
}

fn prove_layout_inheritance(
    source: &OpcPackage,
    source_layout: &PackURI,
    destination: &OpcPackage,
    destination_layout: &PackURI,
    limits: Limits,
) -> Result<()> {
    let mut seen = HashSet::new();
    seen.try_reserve(4).map_err(|source| Error::Allocation {
        resource: "cross-slide inheritance proof identities",
        source,
    })?;
    let relationship_limit = limits
        .max_parts()
        .checked_mul(64)
        .ok_or_else(|| invalid("cross-slide inheritance relationship limit overflow"))?;
    let mut traversed_relationships = 0usize;
    let source_part = source.get_part(source_layout)?;
    let destination_part = destination.get_part(destination_layout)?;
    prove_inherited_part(
        source,
        source_layout,
        destination,
        destination_layout,
        "slide layout",
        InheritanceSurface::Layout,
        &mut seen,
        &mut traversed_relationships,
        relationship_limit,
    )?;
    let source_master = single_inheritance_target(
        source,
        source_part,
        rt::SLIDE_MASTER,
        "slide master",
        ct::PML_SLIDE_MASTER,
    )?;
    let destination_master = single_inheritance_target(
        destination,
        destination_part,
        rt::SLIDE_MASTER,
        "slide master",
        ct::PML_SLIDE_MASTER,
    )?;
    prove_inherited_part(
        source,
        &source_master,
        destination,
        &destination_master,
        "slide master",
        InheritanceSurface::Master,
        &mut seen,
        &mut traversed_relationships,
        relationship_limit,
    )?;
    let source_master_part = source.get_part(&source_master)?;
    let destination_master_part = destination.get_part(&destination_master)?;
    let source_theme = single_inheritance_target(
        source,
        source_master_part,
        rt::THEME,
        "theme",
        ct::OFC_THEME,
    )?;
    let destination_theme = single_inheritance_target(
        destination,
        destination_master_part,
        rt::THEME,
        "theme",
        ct::OFC_THEME,
    )?;
    prove_inherited_part(
        source,
        &source_theme,
        destination,
        &destination_theme,
        "theme",
        InheritanceSurface::Theme,
        &mut seen,
        &mut traversed_relationships,
        relationship_limit,
    )
}

fn prove_inherited_part(
    source: &OpcPackage,
    source_name: &PackURI,
    destination: &OpcPackage,
    destination_name: &PackURI,
    label: &'static str,
    surface: InheritanceSurface,
    seen: &mut HashSet<(PackURI, PackURI)>,
    traversed_relationships: &mut usize,
    relationship_limit: usize,
) -> Result<()> {
    if !seen.insert((source_name.clone(), destination_name.clone())) {
        return Ok(());
    }
    let left = source.get_part(source_name)?;
    let right = destination.get_part(destination_name)?;
    reject_mce(left.blob(), label)?;
    reject_mce(right.blob(), label)?;
    let expected_content_type = inheritance_content_type(surface);
    if left.content_type() != expected_content_type || right.content_type() != expected_content_type
    {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("source and destination {label} parts have incompatible content types"),
        );
    }
    if left.content_type() != right.content_type() || left.blob() != right.blob() {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("source and destination {label} parts are not byte-equivalent"),
        );
    }
    reject_inheritance_edges(left, surface, label)?;
    reject_inheritance_edges(right, surface, label)?;
    let mut left_relationships = Vec::new();
    left_relationships
        .try_reserve_exact(left.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide source inheritance relationships",
            source,
        })?;
    left_relationships.extend(left.rels().iter());
    let mut right_relationships = Vec::new();
    right_relationships
        .try_reserve_exact(right.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide destination inheritance relationships",
            source,
        })?;
    right_relationships.extend(right.rels().iter());
    left_relationships.sort_unstable_by(|a, b| a.r_id().cmp(b.r_id()));
    right_relationships.sort_unstable_by(|a, b| a.r_id().cmp(b.r_id()));
    if left_relationships.len() != right_relationships.len() {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("source and destination {label} relationship graphs differ"),
        );
    }
    for (left_relationship, right_relationship) in
        left_relationships.iter().zip(right_relationships)
    {
        *traversed_relationships = (*traversed_relationships)
            .checked_add(1)
            .ok_or_else(|| invalid("cross-slide inheritance relationship count overflow"))?;
        if *traversed_relationships > relationship_limit {
            return Err(Error::Limit {
                resource: "cross-slide inheritance relationships",
                limit: relationship_limit,
            });
        }
        if left_relationship.r_id() != right_relationship.r_id()
            || left_relationship.reltype() != right_relationship.reltype()
            || left_relationship.is_external() != right_relationship.is_external()
            || left_relationship.target_mode() != right_relationship.target_mode()
            || left_relationship.target_ref() != right_relationship.target_ref()
        {
            return refusal(
                SlideCopyRefusal::SharedOwner,
                format!("source and destination {label} relationships differ"),
            );
        }
        if left_relationship.is_external() {
            if left_relationship.target_ref() != right_relationship.target_ref() {
                return refusal(
                    SlideCopyRefusal::SharedOwner,
                    format!("source and destination {label} external targets differ"),
                );
            }
            continue;
        }
        let left_target = left_relationship.target_partname()?;
        let right_target = right_relationship.target_partname()?;
        let Some(next_surface) = inherited_surface(surface, left_relationship.reltype()) else {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                format!("{label} contains an unsupported internal inheritance edge"),
            );
        };
        let left_part = source.get_part(&left_target)?;
        let right_part = destination.get_part(&right_target)?;
        let expected_content_type = inheritance_content_type(next_surface);
        if left_part.content_type() != expected_content_type
            || right_part.content_type() != expected_content_type
        {
            return refusal(
                SlideCopyRefusal::SharedOwner,
                format!(
                    "source and destination {label} relationship targets have incompatible content types"
                ),
            );
        }
        if left_part.content_type() != right_part.content_type()
            || left_part.blob() != right_part.blob()
        {
            return refusal(
                SlideCopyRefusal::SharedOwner,
                format!("source and destination {label} relationship targets differ"),
            );
        }
        seen.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "cross-slide inheritance proof identities",
            source,
        })?;
        prove_inherited_part(
            source,
            &left_target,
            destination,
            &right_target,
            label_for_surface(next_surface),
            next_surface,
            seen,
            traversed_relationships,
            relationship_limit,
        )?;
    }
    Ok(())
}

fn inherited_surface(
    surface: InheritanceSurface,
    relationship_type: &str,
) -> Option<InheritanceSurface> {
    match surface {
        InheritanceSurface::Layout
            if crate::parts::is_relationship_type(
                relationship_type,
                rt::SLIDE_MASTER,
                "slideMaster",
            ) =>
        {
            Some(InheritanceSurface::Master)
        },
        InheritanceSurface::Master
            if crate::parts::is_relationship_type(
                relationship_type,
                rt::SLIDE_LAYOUT,
                "slideLayout",
            ) =>
        {
            Some(InheritanceSurface::Layout)
        },
        InheritanceSurface::Master
            if crate::parts::is_relationship_type(relationship_type, rt::THEME, "theme") =>
        {
            Some(InheritanceSurface::Theme)
        },
        _ => None,
    }
}

fn label_for_surface(surface: InheritanceSurface) -> &'static str {
    match surface {
        InheritanceSurface::Layout => "slide layout",
        InheritanceSurface::Master => "slide master",
        InheritanceSurface::Theme => "theme",
    }
}

fn inheritance_content_type(surface: InheritanceSurface) -> &'static str {
    match surface {
        InheritanceSurface::Layout => ct::PML_SLIDE_LAYOUT,
        InheritanceSurface::Master => ct::PML_SLIDE_MASTER,
        InheritanceSurface::Theme => ct::OFC_THEME,
    }
}

fn single_inheritance_target(
    package: &OpcPackage,
    part: &dyn Part,
    relationship_type: &str,
    label: &'static str,
    expected_content_type: &'static str,
) -> Result<PackURI> {
    let mut target = None;
    for relationship in part.rels().iter() {
        if !crate::parts::is_relationship_type(relationship.reltype(), relationship_type, label) {
            continue;
        }
        if relationship.is_external()
            || relationship.target_query().is_some()
            || relationship.target_fragment().is_some()
            || relationship.target_mode() != TargetMode::Internal
        {
            return refusal(
                SlideCopyRefusal::SharedOwner,
                format!("the inherited {label} relationship is not an exact internal edge"),
            );
        }
        if target.is_some() {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("an inherited part has more than one {label} relationship"),
            );
        }
        target = Some(relationship.target_partname()?);
    }
    let target = target.ok_or_else(|| Error::SlideCopyPlan {
        kind: SlideCopyRefusal::SharedOwner,
        detail: format!("the inherited part has no {label} relationship"),
    })?;
    let target_part = package.get_part(&target)?;
    if target_part.content_type() != expected_content_type {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("the inherited {label} relationship has an unexpected content type"),
        );
    }
    Ok(target)
}

fn find_slide_by_part(snapshot: &Snapshot, part: &PackURI) -> Result<Slide> {
    snapshot
        .slides
        .iter()
        .find(|slide| slide.part_name == *part)
        .cloned()
        .ok_or_else(|| Error::SlideCopyPlan {
            kind: SlideCopyRefusal::SharedOwner,
            detail: "the durable cross-slide patch source slide is not presentation-owned"
                .to_owned(),
        })
}

fn reject_slide_name_collisions(
    source_slides: &[Slide],
    destination_slides: &[Slide],
    selected_source: &Slide,
) -> Result<()> {
    let mut source_names = HashSet::new();
    source_names
        .try_reserve(source_slides.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide source slide names",
            source,
        })?;
    for slide in source_slides {
        if !source_names.insert(slide.name.as_str()) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "the source presentation contains duplicate producer-visible slide names",
            );
        }
    }
    let mut destination_names = HashSet::new();
    destination_names
        .try_reserve(destination_slides.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide destination slide names",
            source,
        })?;
    for slide in destination_slides {
        if !destination_names.insert(slide.name.as_str()) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "the destination presentation contains duplicate producer-visible slide names",
            );
        }
        if slide.name == selected_source.name {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "the copied slide's producer-visible name collides with a destination slide",
            );
        }
    }
    Ok(())
}

#[cfg(test)]
#[test]
fn mce_retention_intersection_preserves_the_tighter_policy_in_both_orders() -> Result<()> {
    for maximum in [0, 1, 1024, usize::MAX] {
        let left = Limits::default().with_max_retained_mce_bytes(maximum);
        for other in [0, 512, 1024 * 1024, usize::MAX] {
            let right = Limits::default().with_max_retained_mce_bytes(other);
            let expected = Limits::default().with_max_retained_mce_bytes(maximum.min(other));
            assert_eq!(intersect_limits(left, right)?, expected);
            assert_eq!(intersect_limits(right, left)?, expected);
        }
    }
    Ok(())
}

fn intersect_limits(left: Limits, right: Limits) -> Result<Limits> {
    Limits::new(
        left.max_parts().min(right.max_parts()),
        left.max_patch_bytes().min(right.max_patch_bytes()),
        left.max_text_bytes().min(right.max_text_bytes()),
        left.max_history_entries().min(right.max_history_entries()),
        left.max_history_bytes().min(right.max_history_bytes()),
        left.max_retained_candidate_bytes()
            .min(right.max_retained_candidate_bytes()),
    )
    .map(|limits| {
        limits.with_max_retained_mce_bytes(
            left.max_retained_mce_bytes()
                .min(right.max_retained_mce_bytes()),
        )
    })
    .ok_or_else(|| invalid("cross-slide copy limits are invalid"))
}

fn validate_slide_ids(slides: &[Slide]) -> Result<()> {
    if slides
        .iter()
        .any(|slide| !(256..=2_147_483_647).contains(&slide.id))
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "an existing slide ID is outside the PresentationML range",
        );
    }
    Ok(())
}

fn next_slide_id(slides: &[Slide]) -> Result<u32> {
    let mut used = Vec::new();
    used.try_reserve_exact(slides.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide used slide IDs",
            source,
        })?;
    used.extend(slides.iter().map(|slide| slide.id));
    used.sort_unstable();
    let mut candidate = used.last().copied().unwrap_or(255).max(255);
    if candidate < 2_147_483_647 {
        return candidate
            .checked_add(1)
            .ok_or_else(|| invalid("cross-slide slide ID overflow"));
    }
    candidate = 256;
    for value in used {
        if value == candidate {
            candidate = candidate
                .checked_add(1)
                .ok_or_else(|| invalid("cross-slide slide ID overflow"))?;
        } else if value > candidate {
            return Ok(candidate);
        }
    }
    if candidate <= 2_147_483_647 {
        Ok(candidate)
    } else {
        refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "the PresentationML slide-ID space is exhausted",
        )
    }
}

fn next_relationship_id(relationships: &litchi_opc::Relationships) -> Result<String> {
    let mut used = Vec::new();
    used.try_reserve_exact(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "cross-slide used relationship IDs",
            source,
        })?;
    used.extend(relationships.iter().filter_map(|relationship| {
        relationship
            .r_id()
            .strip_prefix("rId")
            .and_then(|value| value.parse::<u32>().ok())
    }));
    used.sort_unstable();
    used.dedup();
    let mut candidate = 1u32;
    for value in used {
        if value == candidate {
            candidate = candidate
                .checked_add(1)
                .ok_or_else(|| invalid("cross-slide relationship-ID space is exhausted"))?;
        } else if value > candidate {
            break;
        }
    }
    Ok(format!("rId{candidate}"))
}

fn reject_protected(xml: &[u8], context: &'static str) -> Result<()> {
    let text = std::str::from_utf8(xml)
        .map_err(|error| Error::Xml(format!("{context} XML is not UTF-8: {error}")))?;
    if crate::presentation_properties::metadata::protection::Settings::parse_xml(text)?
        .is_protected()
    {
        return refusal(
            SlideCopyRefusal::ProtectedPresentation,
            format!("{context} has an active modify-password verifier"),
        );
    }
    Ok(())
}

fn has_macro_infrastructure(package: &OpcPackage) -> bool {
    package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
        )
    }) || package.iter_parts().any(|part| {
        matches!(
            part.content_type(),
            ct::OFC_VBA_PROJECT
                | ct::OFC_VBA_PROJECT_SIGNATURE
                | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
                | ct::PML_PRES_MACRO_MAIN
                | ct::PML_SLIDESHOW_MACRO_MAIN
                | ct::PML_TEMPLATE_MACRO_MAIN
        ) || part.rels().iter().any(|relationship| {
            matches!(
                relationship.reltype(),
                rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
            )
        })
    })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PackageDialect {
    Transitional,
    Strict,
}

fn prove_package_dialect(package: &OpcPackage, presentation: &dyn Part) -> Result<PackageDialect> {
    let (transitional, strict) = dialect_namespace_flags(presentation.blob());
    let dialect = match (transitional, strict) {
        (true, false) => PackageDialect::Transitional,
        (false, true) => PackageDialect::Strict,
        (true, true) => {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                "the presentation owner mixes strict and transitional XML namespaces",
            );
        },
        (false, false) => {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                "the presentation owner does not declare a recognized strict or transitional namespace",
            );
        },
    };

    for part in package.iter_parts() {
        if !is_xml_part(part.partname(), part.content_type()) {
            continue;
        }
        // The dialect check reads every XML part's markup, so every XML part
        // is decoded here (ADR 0030).
        let part_blob = package.get_part(part.partname()).map_err(Error::from)?;
        let (part_transitional, part_strict) = dialect_namespace_flags(part_blob.blob());
        if part_transitional && part_strict {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                "an XML part mixes strict and transitional namespaces",
            );
        }
        let conflicts = match dialect {
            PackageDialect::Transitional => part_strict,
            PackageDialect::Strict => part_transitional,
        };
        if conflicts {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                "an XML dependency uses a dialect different from its presentation package",
            );
        }
        prove_relationship_dialect(part.rels(), dialect)?;
    }
    prove_relationship_dialect(package.rels(), dialect)?;
    Ok(dialect)
}

fn prove_relationship_dialect(
    relationships: &litchi_opc::Relationships,
    dialect: PackageDialect,
) -> Result<()> {
    for relationship in relationships.iter() {
        let Some(actual) = relationship_dialect(relationship.reltype()) else {
            continue;
        };
        if actual != dialect {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                "a package relationship uses a dialect different from its presentation package",
            );
        }
    }
    Ok(())
}

fn relationship_dialect(value: &str) -> Option<PackageDialect> {
    if value.starts_with(TRANSITIONAL_REL) {
        Some(PackageDialect::Transitional)
    } else if value.starts_with(STRICT_REL) {
        Some(PackageDialect::Strict)
    } else {
        None
    }
}

fn dialect_namespace_flags(bytes: &[u8]) -> (bool, bool) {
    let transitional = [
        TRANSITIONAL_PML,
        TRANSITIONAL_DML,
        TRANSITIONAL_CHART,
        TRANSITIONAL_DIAGRAM,
        TRANSITIONAL_CHART_DRAWING,
        TRANSITIONAL_REL_NS,
    ]
    .iter()
    .any(|namespace| {
        bytes
            .windows(namespace.len())
            .any(|window| window == *namespace)
    });
    let strict = [
        STRICT_PML,
        STRICT_DML,
        STRICT_CHART,
        STRICT_DIAGRAM,
        STRICT_CHART_DRAWING,
        STRICT_REL_NS,
    ]
    .iter()
    .any(|namespace| {
        bytes
            .windows(namespace.len())
            .any(|window| window == *namespace)
    });
    (transitional, strict)
}

fn is_xml_part(partname: &PackURI, content_type: &str) -> bool {
    content_type == "application/xml"
        || content_type.ends_with("+xml")
        || partname.membername().ends_with(".xml")
        || partname.membername().ends_with(".rels")
}

/// Hash the exact serialized archive used for physical authorization.
///
/// `OpcPackage::to_stream` is source-aware: an untouched package streams its
/// retained source archive byte-for-byte (including ZIP ordering, compression,
/// comments, and extra fields), while an authored or mutated package streams
/// the complete checked OPC graph. Unknown non-Part members are refused before
/// this helper, so no unmodeled ZIP item can be silently dropped.
fn physical_package_fingerprint(package: &OpcPackage, limits: Limits) -> Result<[u8; 32]> {
    reject_unknown_non_part_members(package, "cross-slide physical authorization")?;
    let mut sink = ArchiveHashWriter::new(limits.max_patch_bytes());
    let result = package.to_stream(&mut sink);
    if sink.exceeded {
        return Err(Error::Limit {
            resource: "cross-slide serialized archive bytes",
            limit: limits.max_patch_bytes(),
        });
    }
    result?;
    seal_physical_revision(sink.digest.finalize().into(), sink.length)
}

/// Bind an archive digest and its length into the physical revision.
///
/// This is the only place the `litchi-pptx-cross-physical-v2` domain string and
/// its length prefix are combined, so every route to a physical revision — the
/// streaming hash sink and the candidate archive that was hashed while it was
/// being built — produces exactly the same value.
fn seal_physical_revision(archive_digest: [u8; 32], length: usize) -> Result<[u8; 32]> {
    let length = u64::try_from(length)
        .map_err(|_error| invalid("cross-slide physical package exceeds u64"))?;
    let mut digest = Sha256::new();
    digest.update(b"litchi-pptx-cross-physical-v2");
    digest.update(length.to_le_bytes());
    digest.update(archive_digest);
    Ok(digest.finalize().into())
}

/// Serialized-archive revision of an immutable snapshot's package.
///
/// A [`Snapshot`] owns its `OpcPackage` behind an `Arc` and never mutates it,
/// so [`physical_package_fingerprint`] is a pure function of the snapshot and
/// the archive bound it is taken under. The first call under a given bound
/// fills the snapshot's cache and every later call returns the same value the
/// recomputation would. The unknown-non-Part-member refusal still runs on every
/// call, so no refusal moves; only the repeated serialization and hash of
/// unchanged bytes is skipped.
fn snapshot_physical_revision(snapshot: &Snapshot, limits: Limits) -> Result<[u8; 32]> {
    let bound = limits.max_patch_bytes();
    if let Some(&(cached_bound, revision)) = snapshot.physical_revision.get()
        && cached_bound == bound
    {
        reject_unknown_non_part_members(
            snapshot.package.as_ref(),
            "cross-slide physical authorization",
        )?;
        debug_assert!(
            physical_package_fingerprint(snapshot.package.as_ref(), limits)
                .is_ok_and(|fresh| fresh == revision),
            "cross-slide reused a stale serialized package revision"
        );
        return Ok(revision);
    }
    let revision = physical_package_fingerprint(snapshot.package.as_ref(), limits)?;
    let _first = snapshot.physical_revision.set((bound, revision));
    Ok(revision)
}

/// Record a serialized-archive revision an application already proved.
///
/// Application hashes the live package before capturing it, and `capture`
/// clones that package into the snapshot. `OpcPackage` clones carry the
/// retained source archive, its exact-source authorization and the complete
/// graph, so the clone serializes to the same bytes as the package the caller
/// hashed. The `debug_assert` re-derives the value on the snapshot's own
/// package in test and debug builds.
fn remember_physical_revision(snapshot: &Snapshot, limits: Limits, revision: [u8; 32]) {
    debug_assert!(
        physical_package_fingerprint(snapshot.package.as_ref(), limits)
            .is_ok_and(|fresh| fresh == revision),
        "cross-slide seeded a snapshot with a foreign serialized package revision"
    );
    let _first = snapshot
        .physical_revision
        .set((limits.max_patch_bytes(), revision));
}

/// Physical revision of a candidate whose archive bytes were already hashed.
///
/// `build_candidate` hashes the candidate archive while it serializes it and
/// then reopens that exact `Vec`, so an exact-source-authorized reopen
/// republishes those bytes verbatim. The unknown-member refusal still runs
/// here, in the position the recomputation ran it.
fn candidate_physical_revision(
    candidate: &OpcPackage,
    limits: Limits,
    known: Option<[u8; 32]>,
) -> Result<[u8; 32]> {
    let Some(known) = known else {
        return physical_package_fingerprint(candidate, limits);
    };
    reject_unknown_non_part_members(candidate, "cross-slide physical authorization")?;
    debug_assert!(
        physical_package_fingerprint(candidate, limits).is_ok_and(|fresh| fresh == known),
        "cross-slide reused a stale candidate archive revision"
    );
    Ok(known)
}

struct ArchiveHashWriter {
    digest: Sha256,
    length: usize,
    limit: usize,
    exceeded: bool,
}

impl ArchiveHashWriter {
    fn new(limit: usize) -> Self {
        Self {
            digest: Sha256::new(),
            length: 0,
            limit,
            exceeded: false,
        }
    }
}

impl Write for ArchiveHashWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let Some(next) = self.length.checked_add(bytes.len()) else {
            self.exceeded = true;
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "cross-slide serialized archive length overflow",
            ));
        };
        if next > self.limit {
            self.exceeded = true;
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "cross-slide serialized archive exceeds its bound",
            ));
        }
        self.digest.update(bytes);
        self.length = next;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct BoundedVecWriter {
    bytes: Vec<u8>,
    digest: Sha256,
    limit: usize,
    exceeded: bool,
    allocation_failure: Option<TryReserveError>,
}

impl BoundedVecWriter {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            digest: Sha256::new(),
            limit,
            exceeded: false,
            allocation_failure: None,
        }
    }

    fn into_bytes(self) -> Result<(Vec<u8>, [u8; 32])> {
        let digest = self.digest.finalize().into();
        if self.bytes.capacity() == self.bytes.len() {
            return Ok((self.bytes, digest));
        }
        // Owned OPC ingress retains the Vec, including spare capacity. Copy
        // once into a fallibly reserved buffer instead of retaining geometric
        // headroom. Both buffers may coexist; the archive limit bounds output
        // length, not aggregate live memory or allocator rounding.
        let mut compact = Vec::new();
        compact
            .try_reserve_exact(self.bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "cross-slide candidate archive",
                source,
            })?;
        compact.extend_from_slice(&self.bytes);
        Ok((compact, digest))
    }
}

impl Write for BoundedVecWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let Some(next) = self.bytes.len().checked_add(bytes.len()) else {
            self.exceeded = true;
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "cross-slide candidate archive length overflow",
            ));
        };
        if next > self.limit {
            self.exceeded = true;
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "cross-slide candidate archive exceeds its bound",
            ));
        }
        if next > self.bytes.capacity() {
            // Repeated small ZIP writes must not request a new allocation for
            // every chunk. Keep deliberate growth within the archive bound;
            // the checked length above guarantees target >= len.
            let target = self
                .bytes
                .capacity()
                .saturating_mul(2)
                .max(next)
                .min(self.limit);
            if let Err(source) = self.bytes.try_reserve_exact(target - self.bytes.len()) {
                self.allocation_failure = Some(source);
                return Err(io::Error::other(
                    "cross-slide candidate archive allocation failed",
                ));
            }
        }
        self.digest.update(bytes);
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Serialize `package` into a bounded owned archive and return its digest.
///
/// The digest is taken over exactly the bytes the writer accepted, which is
/// the same stream [`ArchiveHashWriter`] would have hashed, so the caller can
/// seal it into a physical revision instead of serializing the package a
/// second time.
fn bounded_package_bytes(package: &OpcPackage, limit: usize) -> Result<(Vec<u8>, [u8; 32])> {
    let mut writer = BoundedVecWriter::new(limit);
    let result = package.to_stream(&mut writer);
    if let Some(source) = writer.allocation_failure.take() {
        return Err(Error::Allocation {
            resource: "cross-slide candidate archive",
            source,
        });
    }
    if writer.exceeded {
        return Err(Error::Limit {
            resource: "cross-slide serialized archive bytes",
            limit,
        });
    }
    result?;
    writer.into_bytes()
}

fn copy_string(value: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

fn validate_patch_descriptor(
    patch: &Patch,
    source_slide: &PackURI,
    destination_slide: &PackURI,
    destination_layout: &PackURI,
    _position: usize,
    slide_id: u32,
    presentation_relationship_id: &str,
) -> Result<()> {
    if source_slide.as_str().is_empty()
        || destination_slide.as_str().is_empty()
        || destination_layout.as_str().is_empty()
    {
        return Err(invalid(
            "cross-slide durable patch has invalid slide identity",
        ));
    }
    if !(256..=2_147_483_647).contains(&slide_id) || presentation_relationship_id.is_empty() {
        return Err(invalid(
            "cross-slide durable patch has invalid destination identity",
        ));
    }
    if patch.is_empty() {
        return Err(invalid(
            "cross-slide durable patch has no destination change",
        ));
    }
    Ok(())
}

fn put_text(output: &mut Vec<u8>, value: &str, field: &'static str) -> Result<()> {
    let length = u32::try_from(value.len())
        .map_err(|_error| invalid(format!("cross-slide {field} exceeds u32")))?;
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value.as_bytes());
    Ok(())
}

fn put_u32(output: &mut Vec<u8>, value: u32) -> Result<()> {
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn put_u64(output: &mut Vec<u8>, value: u64) -> Result<()> {
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn parse_part_name_text(value: String) -> Result<PackURI> {
    PackURI::new(value).map_err(Error::Invalid)
}

fn refusal<T>(kind: SlideCopyRefusal, detail: impl Into<String>) -> Result<T> {
    Err(Error::SlideCopyPlan {
        kind,
        detail: detail.into(),
    })
}

struct WireInput<'a> {
    bytes: &'a [u8],
    position: usize,
}

impl<'a> WireInput<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, position: 0 }
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| invalid("cross-slide durable patch position overflow"))?;
        let value = self
            .bytes
            .get(self.position..end)
            .ok_or_else(|| invalid("cross-slide durable patch is truncated"))?;
        self.position = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8> {
        self.take(1)?
            .first()
            .copied()
            .ok_or_else(|| invalid("cross-slide durable patch u8 is malformed"))
    }

    fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().map_err(
            |_error| invalid("cross-slide durable patch u32 is malformed"),
        )?))
    }

    fn usize64(&mut self, field: &'static str) -> Result<usize> {
        usize::try_from(u64::from_le_bytes(self.take(8)?.try_into().map_err(
            |_error| invalid(format!("cross-slide durable patch {field} is malformed")),
        )?))
        .map_err(|_error| invalid(format!("cross-slide durable patch {field} exceeds usize")))
    }

    fn revision(&mut self) -> Result<[u8; 32]> {
        self.take(32)?
            .try_into()
            .map_err(|_error| invalid("cross-slide durable patch revision is malformed"))
    }

    fn text32(&mut self, field: &'static str) -> Result<String> {
        let length = usize::try_from(self.u32()?)
            .map_err(|_error| invalid(format!("cross-slide {field} length exceeds usize")))?;
        let value = std::str::from_utf8(self.take(length)?)
            .map_err(|error| invalid(format!("cross-slide {field} is not UTF-8: {error}")))?;
        let mut output = String::new();
        output
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "cross-slide durable patch text",
                source,
            })?;
        output.push_str(value);
        Ok(output)
    }

    fn is_empty(&self) -> bool {
        self.position == self.bytes.len()
    }
}

#[cfg(test)]
mod bounded_writer_tests;
#[cfg(test)]
mod media_transfer_tests;
#[cfg(test)]
mod retention_tests;
#[cfg(test)]
mod revision_cache_tests;
