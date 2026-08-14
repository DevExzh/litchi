//! Bounded cross-presentation slide-copy planning.
//!
//! A cross-package copy is deliberately more restrictive than a same-package
//! copy.  The source slide's private owned-role closure is copied byte-for-byte
//! into the destination, while the source layout is never copied: it is reused
//! only after the layout/master/theme inheritance graph has been proven
//! equivalent to the destination slide's selected layout.  The resulting
//! destination graph is captured as an exact, complete-revision-bound patch.

use std::collections::{HashMap, HashSet, TryReserveError};
use std::io::{self, Write};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use sha2::{Digest, Sha256};

use super::copy_plan::{
    available_copy_name_avoiding, collect_owned_closure, has_signature_infrastructure, reject_mce,
    reject_unknown_non_part_members, resolve_slide, validate_registered_layout,
    validate_slide_surface,
};
use super::model::{Limits, Slide, Snapshot, invalid};
use super::patch::Patch;
use crate::{Error, Result, SlideCopyRefusal};

const PATCH_MAGIC: &[u8; 8] = b"LPCP0002";
const PATCH_HEADER_BYTES: usize = PATCH_MAGIC.len() + (6 * 32) + (4 * 4) + 8 + 4 + 8;
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

/// An immutable plan to copy one source slide into a different destination
/// presentation.
///
/// The source slide's raw bytes and supported owned dependency closure are
/// retained exactly.  The destination's selected slide supplies the layout
/// boundary; the source layout, master, and theme are never copied or
/// rewritten.  Planning validates a complete candidate before returning.
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
}

impl CrossSlideCopyPlan {
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
    patch: Patch,
}

impl CrossSlideCopyPatch {
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
            patch: self.patch.inverse(),
        }
    }

    /// Serialize this patch into the stable `LPCP0002` binary format.
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
        let payload_len = u64::try_from(payload.len())
            .map_err(|_error| invalid("cross-slide durable patch payload exceeds u64"))?;
        put_u64(&mut output, payload_len)?;
        output.extend_from_slice(&payload);
        Ok(output)
    }

    /// Parse a stable durable patch under conservative finite limits.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a stable durable patch under caller-selected finite limits.
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
    pub fn plan_cross_slide_copy<'s, 'd>(
        &self,
        source: &Snapshot,
        source_slide: impl Into<crate::slide::Key<'s>>,
        destination_slide: impl Into<crate::slide::Key<'d>>,
        position: usize,
    ) -> Result<CrossSlideCopyPlan> {
        let source_slide = resolve_slide(source, source_slide.into())?;
        let destination_slide = resolve_slide(self, destination_slide.into())?;
        plan_cross_slide_copy_for_slides(source, self, source_slide, destination_slide, position)
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
    if super::model::package_fingerprint(source)? != plan.source_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the complete source package graph changed after cross-slide planning",
        });
    }
    if super::model::package_fingerprint(destination)? != plan.destination_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the complete destination package graph changed after cross-slide planning",
        });
    }
    if physical_package_fingerprint(source, plan.patch.patch.limits())?
        != plan.source_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the serialized source package changed after cross-slide planning",
        });
    }
    if physical_package_fingerprint(destination, plan.patch.patch.limits())?
        != plan.destination_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the serialized destination package changed after cross-slide planning",
        });
    }
    let source_snapshot = super::model::capture(
        source,
        plan.patch.patch.limits(),
        source_physical_source_provenance,
    )?;
    let destination_snapshot = super::model::capture(
        destination,
        plan.patch.patch.limits(),
        destination_physical_source_provenance,
    )?;
    let fresh = plan_cross_slide_copy_for_slides(
        &source_snapshot,
        &destination_snapshot,
        plan.source.clone(),
        plan.destination.clone(),
        plan.position,
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
    let mut candidate = destination.clone();
    let snapshot = super::patch::apply_exact_revision(
        &mut candidate,
        &plan.patch.patch,
        plan.target_revision,
        destination_physical_source_provenance,
    )?;
    if physical_package_fingerprint(&candidate, plan.patch.patch.limits())?
        != plan.target_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the published candidate has an unexpected serialized package revision",
        });
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
    if super::model::package_fingerprint(source)? != patch.source_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the complete source package graph differs from the cross-slide patch source",
        });
    }
    if super::model::package_fingerprint(destination)? != patch.destination_revision {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the complete destination package graph differs from the cross-slide patch source",
        });
    }
    if physical_package_fingerprint(source, patch.patch.limits())? != patch.source_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the serialized source package differs from the cross-slide patch source",
        });
    }
    if physical_package_fingerprint(destination, patch.patch.limits())?
        != patch.destination_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the serialized destination package differs from the cross-slide patch source",
        });
    }
    let source_snapshot = super::model::capture(
        source,
        patch.patch.limits(),
        source_physical_source_provenance,
    )?;
    let destination_snapshot = super::model::capture(
        destination,
        patch.patch.limits(),
        destination_physical_source_provenance,
    )?;
    let source_slide = find_slide_by_part(&source_snapshot, &patch.source_slide)?;
    let destination_slide = find_slide_by_part(&destination_snapshot, &patch.destination_slide)?;
    let forward_matches = plan_cross_slide_copy_for_slides(
        &source_snapshot,
        &destination_snapshot,
        source_slide.clone(),
        destination_slide.clone(),
        patch.position,
    )
    .map(|fresh| {
        fresh.patch == *patch
            && fresh.target_revision == patch.target_revision
            && fresh.target_physical_revision == patch.target_physical_revision
            && fresh.slide_id == patch.slide_id
            && fresh.presentation_relationship_id == patch.presentation_relationship_id
    })
    .unwrap_or(false);
    let inverse_matches = if forward_matches {
        true
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
            false
        } else if physical_package_fingerprint(&restored, patch.patch.limits()).ok()
            != Some(patch.target_physical_revision)
        {
            false
        } else {
            let restored_snapshot = super::model::capture(
                &restored,
                patch.patch.limits(),
                destination_physical_source_provenance,
            );
            restored_snapshot
                .ok()
                .and_then(|base| {
                    let restored_destination =
                        find_slide_by_part(&base, &patch.destination_slide).ok()?;
                    let forward = plan_cross_slide_copy_for_slides(
                        &source_snapshot,
                        &base,
                        source_slide,
                        restored_destination,
                        patch.position,
                    )
                    .ok()?;
                    Some(forward.patch.inverse() == *patch)
                })
                .unwrap_or(false)
        }
    };
    if !forward_matches && !inverse_matches {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the durable cross-slide patch does not match a freshly proven candidate",
        });
    }
    let mut candidate = destination.clone();
    let snapshot = super::patch::apply_exact_revision(
        &mut candidate,
        &patch.patch,
        patch.target_revision,
        destination_physical_source_provenance,
    )?;
    if physical_package_fingerprint(&candidate, patch.patch.limits())?
        != patch.target_physical_revision
    {
        return Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_patch",
            reason: "the published candidate has an unexpected serialized package revision",
        });
    }
    *destination = candidate;
    Ok(snapshot)
}

fn plan_cross_slide_copy_for_slides(
    source: &Snapshot,
    destination: &Snapshot,
    source_slide: Slide,
    destination_slide: Slide,
    position: usize,
) -> Result<CrossSlideCopyPlan> {
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
    preflight_parts(destination, &parts, planned_bytes)?;
    let source_physical_revision = physical_package_fingerprint(source.package.as_ref(), limits)?;
    let destination_physical_revision =
        physical_package_fingerprint(destination.package.as_ref(), limits)?;
    let candidate = build_candidate(
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
        limits.max_patch_bytes(),
    )?;
    let patch = Patch::capture(
        destination.package.as_ref(),
        &candidate,
        destination.presentation_name.clone(),
        limits,
    )?;
    let target_revision = super::model::package_fingerprint(&candidate)?;
    let target_physical_revision = physical_package_fingerprint(&candidate, limits)?;
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
    Ok(CrossSlideCopyPlan {
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
    })
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
    archive_limit: usize,
) -> Result<OpcPackage> {
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
    let mut candidate = destination.package.as_ref().clone();
    for planned in parts {
        let original = source.package.get_part(&planned.source)?;
        let mut copied = BlobPart::new_shared(
            planned.target.clone(),
            planned.content_type.clone(),
            original.blob_arc(),
        );
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
    let presentation = destination
        .package
        .get_part(&destination.presentation_name)?;
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
    let serialized = bounded_package_bytes(&candidate, archive_limit)?;
    let reopened = OpcPackage::from_vec(serialized)?;
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
    Ok(reopened)
}

fn preflight_parts(
    destination: &Snapshot,
    parts: &[super::SlideCopyPart],
    planned_bytes: usize,
) -> Result<()> {
    let owner = destination
        .package
        .get_part(&destination.presentation_name)?;
    let estimate = planned_bytes
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
        .ok_or_else(|| invalid("cross-slide candidate byte count overflow"))?;
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

fn intersect_limits(left: Limits, right: Limits) -> Result<Limits> {
    Limits::new(
        left.max_parts().min(right.max_parts()),
        left.max_patch_bytes().min(right.max_patch_bytes()),
        left.max_text_bytes().min(right.max_text_bytes()),
        left.max_history_entries().min(right.max_history_entries()),
        left.max_history_bytes().min(right.max_history_bytes()),
    )
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
        if !is_xml_part(part) {
            continue;
        }
        let (part_transitional, part_strict) = dialect_namespace_flags(part.blob());
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

fn is_xml_part(part: &dyn Part) -> bool {
    let content_type = part.content_type();
    content_type == "application/xml"
        || content_type.ends_with("+xml")
        || part.partname().membername().ends_with(".xml")
        || part.partname().membername().ends_with(".rels")
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
    let length = u64::try_from(sink.length)
        .map_err(|_error| invalid("cross-slide physical package exceeds u64"))?;
    let mut digest = Sha256::new();
    digest.update(b"litchi-pptx-cross-physical-v2");
    digest.update(length.to_le_bytes());
    digest.update(sink.digest.finalize());
    Ok(digest.finalize().into())
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
    limit: usize,
    exceeded: bool,
    allocation_failure: Option<TryReserveError>,
}

impl BoundedVecWriter {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
            exceeded: false,
            allocation_failure: None,
        }
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
        if let Err(source) = self.bytes.try_reserve_exact(bytes.len()) {
            self.allocation_failure = Some(source);
            return Err(io::Error::other(
                "cross-slide candidate archive allocation failed",
            ));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn bounded_package_bytes(package: &OpcPackage, limit: usize) -> Result<Vec<u8>> {
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
    Ok(writer.bytes)
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
