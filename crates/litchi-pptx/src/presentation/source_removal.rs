//! Source-native planning for a bounded slide removal. The slide leaves with
//! the Parts it owns exclusively (record 0303).
//!
//! The plan in this module retains a source-authorized symbolic removal
//! description. It is therefore reusable by the primary source editor, but it
//! deliberately does not model a durable or inverse edit and does not mutate
//! the primary source-backed presentation editor.

#[cfg(test)]
use std::io::Write;
use std::sync::Arc;

use litchi_core::{Position, Reservation, Resource, SourceVersion};
#[cfg(test)]
use litchi_opc::PreparedTopology;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::xml_splice::authored_xml_requires_source_proof;
use litchi_opc::{
    AuthoredXmlFragment, PackURI, Relationship, SourceBackedPackage, SourceLimits,
    SourceTopologyPlan,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::source::SourceBackedPresentationEditor;
use crate::opened::{validate_removal_presentation_owner, validate_removal_slide_surface};
use crate::{Error, Result, SlideRemovalRefusal};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// A checked source-native whole-slide removal prerequisite.
///
/// The topology removes one presentation relationship, one slide Part and
/// the Parts that slide owns exclusively (record 0303). Each
/// raw candidate keeps untouched members as lazy source transfers. The plan is
/// process-local and one-way: it carries no durable revision patch and no
/// inverse operation.
pub(crate) struct SourceBackedSlideRemovalPlan {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    package: SourceBackedPackage,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    description: RemovalDescription,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    source_version: SourceVersion,
    _candidate_xml_reservation: Option<Arc<Reservation>>,
}

/// Symbolic removal metadata retained beside the source package.
///
/// Keeping this description lets the primary source editor rebuild the typed
/// topology later without treating a prepared candidate as a durable patch.
#[derive(Clone)]
struct RemovalDescription {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    presentation_uri: PackURI,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    presentation_relationship_id: String,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    candidate_presentation_xml: Arc<Vec<u8>>,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    source_xml: Option<litchi_opc::SourceXmlPart>,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    removed_slide: PackURI,
    /// Parts the removed slide owned exclusively (record 0303).
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the source-level engine tests")
    )]
    removed_dependents: Vec<PackURI>,
}

impl SourceBackedSlideRemovalPlan {
    /// Prepare and return the checked OPC topology candidate through an
    /// explicit raw view.
    #[cfg(test)]
    pub(crate) fn raw(&self) -> Result<PreparedTopology> {
        self.package.check_execution()?;
        if self.package.source_version()? != self.source_version {
            return Err(Error::PatchConflict);
        }
        let topology = self.rebuild_topology()?;
        let prepared = self.package.clone().prepare_topology(topology)?;
        self.package.check_execution()?;
        if self.package.source_version()? != self.source_version {
            return Err(Error::PatchConflict);
        }
        Ok(prepared)
    }

    /// Publish the checked candidate to a sequential writer.
    ///
    /// The underlying prepared topology checks the source lineage, source
    /// revision, execution policy, and all retained transfer sources again
    /// before accepting output.
    #[cfg(test)]
    pub(crate) fn write_to<W: Write>(&self, writer: W) -> Result<()> {
        Ok(self.raw()?.publish_to_stream(writer)?)
    }

    /// Rebuild the typed source topology represented by this process-local
    /// description for the next primary integration layer.
    #[cfg(test)]
    pub(super) fn rebuild_topology(&self) -> Result<SourceTopologyPlan> {
        let mut topology = SourceTopologyPlan::new();
        if let Some(source_xml) = &self.description.source_xml {
            topology.try_replace_source_xml_part(
                self.description.presentation_uri.clone(),
                source_xml.clone(),
            )?;
        } else {
            topology.try_replace_part_shared(
                self.description.presentation_uri.clone(),
                Arc::clone(&self.description.candidate_presentation_xml),
            )?;
        }
        topology.try_remove_relationship(
            self.description.presentation_uri.clone(),
            self.description.presentation_relationship_id.clone(),
        )?;
        topology.try_remove_part(self.description.removed_slide.clone())?;
        for dependent in &self.description.removed_dependents {
            topology.try_remove_part(dependent.clone())?;
        }
        Ok(topology)
    }
}

impl SourceBackedPresentationEditor {
    /// Plan removal of one slide without mutating this editor.
    ///
    /// The selected slide must be one of at least two slides, have exactly one
    /// internal registered slide-layout relationship, and no incoming owner
    /// relationship besides its presentation binding and the Parts it owns.
    /// The Parts it owns exclusively (notes slide, pictures, diagrams, VML
    /// drawings, comments, embedded objects) are removed with it; Parts that
    /// another owner still targets stay (record 0303). Signature, macro, protected,
    /// MCE, unknown-physical-member, positional, and unknown semantic surfaces
    /// are refused through the typed [`SlideRemovalRefusal`] boundary.
    pub(crate) fn plan_slide_removal(
        &self,
        position: impl Into<Position>,
    ) -> Result<SourceBackedSlideRemovalPlan> {
        self.plan_slide_removal_with_policy(position.into(), false, false)
    }

    /// Re-prove an insertion inverse, including the transition from one
    /// effective slide back to an originally empty catalog.  This is kept
    /// crate-private so the public source removal API continues to refuse a
    /// direct final-slide removal.
    pub(crate) fn plan_slide_removal_for_inverse(
        &self,
        position: impl Into<Position>,
    ) -> Result<SourceBackedSlideRemovalPlan> {
        self.plan_slide_removal_with_policy(position.into(), true, true)
    }

    /// Re-prove a durable patch removal while preserving unrelated opaque
    /// relationship targets. The final-slide policy remains strict; callers
    /// use [`Self::plan_slide_removal_for_inverse`] when the witnessed target
    /// state is intentionally empty.
    pub(crate) fn plan_slide_removal_for_patch(
        &self,
        position: impl Into<Position>,
    ) -> Result<SourceBackedSlideRemovalPlan> {
        self.plan_slide_removal_with_policy(position.into(), false, true)
    }

    fn plan_slide_removal_with_policy(
        &self,
        position: Position,
        allow_empty_result: bool,
        allow_unrelated_opaque_targets: bool,
    ) -> Result<SourceBackedSlideRemovalPlan> {
        let position = position.get();
        self.package.check_execution()?;
        let source_version = self.package.source_version()?;
        let source_slide_count = self.slides.len();
        if position >= source_slide_count {
            return Err(Error::SlideIndexOutOfBounds {
                index: position,
                len: source_slide_count,
            });
        }
        if source_slide_count < 2 && !(allow_empty_result && source_slide_count == 1) {
            return refusal(
                SlideRemovalRefusal::FinalSlide,
                "a presentation must retain at least one slide",
            );
        }

        self.package.validate_topology_source_boundary()?;
        let presentation = self.package.main_document_part()?;
        if !presentation
            .partname()
            .is_equivalent_to(self._presentation.source_partname())
        {
            return Err(Error::PatchConflict);
        }
        validate_view_content_type(&presentation, ct::PML_PRESENTATION_MAIN)?;
        let presentation_xml = self._presentation.source_data().as_bytes();
        if presentation.data()?.as_bytes() != presentation_xml {
            return Err(Error::PatchConflict);
        }
        // The root bytes are already retained by the source facade, but apply
        // the physical source limit before any semantic or protection parser.
        crate::source_limits::validate_xml_limits(presentation_xml, self.limits)?;
        reject_package_features(&self.package, presentation_xml)?;
        map_order_refusal(super::order::validate_source_order_publication(
            &self.package,
            presentation.partname(),
        ))?;
        reject_mce(presentation_xml, &self.package)?;
        validate_removal_presentation_owner(
            presentation_xml,
            self.slides[position].binding.slide_reference_id.as_str(),
        )?;

        let selected = &self.slides[position];
        let relationship_id = selected.binding.slide_reference_id.as_str();
        let relationship = presentation
            .rels()
            .get(relationship_id)
            .ok_or_else(|| Error::Relationship("selected slide relationship is missing".into()))?;
        validate_selected_presentation_relationship(
            relationship,
            relationship_id,
            &selected.part_uri,
        )?;

        let slide = self.package.part(&selected.part_uri)?;
        validate_view_content_type(&slide, ct::PML_SLIDE)?;
        let layout = dependency_free_layout(&self.package, &slide)?;
        map_copy_refusal(
            super::source_cross_copy::validate_registered_removal_layout(
                &self.package,
                &presentation,
                presentation_xml,
                &layout,
                self.package.execution_context().as_ref(),
            ),
        )?;
        let slide_data = slide.data()?;
        crate::source_limits::validate_xml_limits(slide_data.as_bytes(), self.limits)?;
        validate_removal_slide_surface(slide_data.as_bytes())?;

        // Removal publication needs the exact source images for the Part and
        // metadata that an inverse may restore after the forward ZIP has been
        // reopened. These reads are bounded and happen before any candidate
        // topology is admitted or published.
        // The source catalog may intentionally skip a repaired relationship
        // subtree. Removal requires the complete typed relationship graph, so
        // do not infer unique ownership from a non-clean ingress report.
        if !self.package.tolerance_report().is_clean() {
            return refusal(
                SlideRemovalRefusal::UnknownSemanticSurface,
                "the source relationship or content-type manifests required ingress repairs",
            );
        }
        let owned = if crate::presentation::owned_parts::may_own(slide.rels()) {
            crate::presentation::owned_parts::Graph::from_source(
                &self.package,
                relationship_scan_limit(self.limits)?,
            )?
            .owned_by(&selected.part_uri)?
        } else {
            crate::presentation::owned_parts::Owned::none()
        };
        validate_unique_incoming_owner(
            &self.package,
            presentation.partname(),
            &selected.part_uri,
            relationship_id,
            self.limits,
            allow_unrelated_opaque_targets,
            &owned,
        )?;

        self.package.check_execution()?;
        let candidate_xml_reservation = reserve_candidate_xml(
            &self.package,
            presentation_xml.len(),
            presentation.partname().as_str().len(),
            relationship_id.len(),
            selected.part_uri.as_str().len(),
        )?;
        let candidate_xml = crate::opened::remove_slide_binding(
            presentation_xml,
            self.slides
                .iter()
                .map(|slide| (slide.slide_id, slide.binding.slide_reference_id.as_str())),
            selected.slide_id,
            self.package.execution_context().as_ref(),
        )?;
        let candidate_xml = Arc::new(candidate_xml);
        let deletions = crate::opened::slide_binding_removal_spans(
            presentation_xml,
            selected.slide_id,
            self.package.execution_context().as_ref(),
        )?;
        let mut topology = SourceTopologyPlan::new();
        let source_xml = stage_presentation_replacement(
            &presentation,
            presentation_xml,
            &candidate_xml,
            &deletions,
            &mut topology,
        )?;
        topology
            .try_remove_relationship(presentation.partname().clone(), relationship_id.to_owned())?;
        topology.try_remove_part(selected.part_uri.clone())?;
        for dependent in owned.parts() {
            topology.try_remove_part(dependent.clone())?;
        }

        self.package.check_execution()?;
        // Run the complete OPC admission phase before publishing the plan.
        // The symbolic description below is retained so the primary editor
        // can rebuild this topology after composing additional overlays.
        drop(self.package.clone().prepare_topology(topology)?);
        self.package.check_execution()?;
        if self.package.source_version()? != source_version {
            return Err(Error::PatchConflict);
        }
        Ok(SourceBackedSlideRemovalPlan {
            package: self.package.clone(),
            source_version,
            description: RemovalDescription {
                presentation_uri: presentation.partname().clone(),
                presentation_relationship_id: relationship_id.to_owned(),
                candidate_presentation_xml: candidate_xml,
                source_xml,
                removed_slide: selected.part_uri.clone(),
                removed_dependents: owned.parts().to_vec(),
            },
            _candidate_xml_reservation: candidate_xml_reservation,
        })
    }
}

/// Shared policy admission for a primary effective topology. The primary
/// owner applies its final-slide guard to the effective catalog separately.
pub(super) fn validate_primary_root(
    package: &SourceBackedPackage,
    presentation: &PackURI,
    root: &[u8],
    relationship_id: &str,
) -> Result<()> {
    package.validate_topology_source_boundary()?;
    if !package.tolerance_report().is_clean() {
        return refusal(
            SlideRemovalRefusal::UnknownSemanticSurface,
            "the source manifests required ingress repairs",
        );
    }
    crate::source_limits::validate_xml_limits(root, package.read_limits())?;
    reject_package_features(package, root)?;
    map_order_refusal(super::order::validate_source_order_publication(
        package,
        presentation,
    ))?;
    reject_mce(root, package)?;
    validate_removal_presentation_owner(root, relationship_id)
}

fn stage_presentation_replacement(
    presentation: &litchi_opc::PartView<'_>,
    source_bytes: &[u8],
    candidate_bytes: &Arc<Vec<u8>>,
    deletions: &[std::ops::Range<usize>],
    topology: &mut SourceTopologyPlan,
) -> Result<Option<litchi_opc::SourceXmlPart>> {
    if candidate_bytes.as_slice() == source_bytes {
        return Err(Error::Invalid(
            "slide-removal root replacement did not change the source".into(),
        ));
    }
    let requires_proof = authored_xml_requires_source_proof(
        presentation.partname(),
        presentation.content_type(),
        candidate_bytes.as_slice(),
    )?;
    if !requires_proof {
        topology.try_replace_part_shared(
            presentation.partname().clone(),
            Arc::clone(candidate_bytes),
        )?;
        return Ok(None);
    }

    let source_xml = presentation.source_xml()?;
    if source_xml.bytes() != source_bytes {
        return Err(Error::PatchConflict);
    }
    validate_deletions(source_bytes, candidate_bytes.as_slice(), deletions)?;
    let mut proofs = Vec::new();
    for range in deletions {
        let expected = source_bytes
            .get(range.clone())
            .ok_or_else(|| Error::Invalid("slide-removal XML deletion range is invalid".into()))?;
        proofs.push(source_xml.checked_range(range.clone(), expected)?);
    }
    let mut publication = source_xml.into_publication()?;
    for proof in proofs {
        publication.replace(proof, AuthoredXmlFragment::empty())?;
    }
    let staged = publication.finish()?;
    if staged.bytes() != candidate_bytes.as_slice() {
        return Err(Error::Invalid(
            "slide-removal source XML proof changed the candidate bytes".into(),
        ));
    }
    topology.try_replace_source_xml_part(presentation.partname().clone(), staged.clone())?;
    Ok(Some(staged))
}

fn reserve_candidate_xml(
    package: &SourceBackedPackage,
    source_bytes: usize,
    presentation_uri_bytes: usize,
    relationship_id_bytes: usize,
    slide_uri_bytes: usize,
) -> Result<Option<Arc<Reservation>>> {
    // Hold a conservative envelope for the source-sized candidate, symbolic
    // URI/relationship metadata, and the bounded topology bookkeeping until
    // the plan is dropped. This reservation is acquired before the lexical
    // removal helper allocates its output.
    const FIXED_METADATA_BYTES: usize = 512;
    let metadata_bytes = presentation_uri_bytes
        .checked_add(relationship_id_bytes)
        .and_then(|bytes| bytes.checked_add(slide_uri_bytes))
        .and_then(|bytes| bytes.checked_add(FIXED_METADATA_BYTES))
        .ok_or_else(|| Error::Invalid("slide-removal reservation size overflows usize".into()))?;
    let bytes = source_bytes
        .checked_add(metadata_bytes)
        .ok_or_else(|| Error::Invalid("slide-removal reservation size overflows usize".into()))?;
    if bytes == 0 {
        return Ok(None);
    }
    let Some(context) = package.execution_context() else {
        return Ok(None);
    };
    let bytes = u64::try_from(bytes)
        .map_err(|_| Error::Invalid("slide-removal candidate XML size exceeds u64".into()))?;
    context
        .reserve(Resource::Memory, bytes)
        .map(Arc::new)
        .map(Some)
        .map_err(|error| Error::from(litchi_opc::OpcError::Execution(error)))
}

/// The candidate must be the source with exactly these sorted, disjoint,
/// nonempty spans deleted: the removed slide's catalog entry and its section
/// entry (record 0303).
fn validate_deletions(
    source: &[u8],
    candidate: &[u8],
    deletions: &[std::ops::Range<usize>],
) -> Result<()> {
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-removal deletion proof",
            source,
        })?;
    let mut cursor = 0usize;
    for range in deletions {
        if range.start < cursor || range.start >= range.end || range.end > source.len() {
            return Err(Error::Invalid(
                "slide-removal root candidate is not a set of lexical deletions".into(),
            ));
        }
        expected.extend_from_slice(&source[cursor..range.start]);
        cursor = range.end;
    }
    expected.extend_from_slice(&source[cursor..]);
    if deletions.is_empty() || expected != candidate {
        return Err(Error::Invalid(
            "slide-removal root candidate is not a set of lexical deletions".into(),
        ));
    }
    Ok(())
}

fn validate_selected_presentation_relationship(
    relationship: &Relationship,
    relationship_id: &str,
    selected: &PackURI,
) -> Result<()> {
    if relationship.r_id() != relationship_id || relationship.is_external() {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected presentation relationship is not one internal binding",
        );
    }
    if !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide") {
        return refusal(
            SlideRemovalRefusal::UnsupportedRelationship,
            "the selected presentation relationship is not a slide relationship",
        );
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected presentation relationship contains a query or fragment",
        );
    }
    if !relationship.target_partname()?.is_equivalent_to(selected) {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected presentation relationship targets a different slide",
        );
    }
    Ok(())
}

fn dependency_free_layout(
    package: &SourceBackedPackage,
    slide: &litchi_opc::PartView<'_>,
) -> Result<PackURI> {
    let relationship = crate::opened::removal_layout_edge(slide.rels())?;
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected layout target contains a query or fragment",
        );
    }
    let layout = relationship.target_partname()?;
    if !layout.as_str().starts_with("/ppt/slideLayouts/") {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected layout is outside the PresentationML layout owner",
        );
    }
    let layout_part = package.part(&layout)?;
    validate_view_content_type(&layout_part, ct::PML_SLIDE_LAYOUT)?;
    Ok(layout)
}

fn validate_view_content_type(view: &litchi_opc::PartView<'_>, expected: &str) -> Result<()> {
    if view.content_type() == expected {
        return Ok(());
    }
    Err(Error::ContentType {
        expected: expected.to_owned(),
        actual: view.content_type().to_owned(),
    })
}

fn validate_unique_incoming_owner(
    package: &SourceBackedPackage,
    presentation: &PackURI,
    selected: &PackURI,
    selected_relationship_id: &str,
    limits: SourceLimits,
    allow_unrelated_opaque_targets: bool,
    owned: &crate::presentation::owned_parts::Owned,
) -> Result<()> {
    let maximum = relationship_scan_limit(limits)?;
    let mut scanned = 0usize;
    let mut selected_owners = 0usize;
    for relationship in package.rels().iter() {
        scanned = checked_relationship_count(scanned, maximum, package)?;
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        if !target.is_equivalent_to(selected) {
            if !allow_unrelated_opaque_targets
                && (relationship.target_query().is_some()
                    || relationship.target_fragment().is_some())
            {
                return refusal(
                    SlideRemovalRefusal::AmbiguousTopology,
                    "an internal package-root relationship has a query or fragment",
                );
            }
            continue;
        }
        if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
            return refusal(
                SlideRemovalRefusal::AmbiguousTopology,
                "an internal package-root relationship targets the selected slide with a query or fragment",
            );
        }
        return refusal(
            SlideRemovalRefusal::SharedOwner,
            "the package root directly references the selected slide",
        );
    }
    for owner in package.iter_parts() {
        // The Parts the slide owns leave with it, so their edges (a notes
        // slide names its slide) are not owners that survive the removal.
        let leaves = owned.contains(owner.partname());
        for relationship in owner.rels().iter() {
            scanned = checked_relationship_count(scanned, maximum, package)?;
            if leaves {
                continue;
            }
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            if !target.is_equivalent_to(selected) {
                if !allow_unrelated_opaque_targets
                    && (relationship.target_query().is_some()
                        || relationship.target_fragment().is_some())
                {
                    return refusal(
                        SlideRemovalRefusal::AmbiguousTopology,
                        "an internal package relationship has a query or fragment",
                    );
                }
                continue;
            }
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return refusal(
                    SlideRemovalRefusal::AmbiguousTopology,
                    "an internal package relationship targets the selected slide with a query or fragment",
                );
            }
            if owner.partname().is_equivalent_to(presentation)
                && relationship.r_id() == selected_relationship_id
                && crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
            {
                selected_owners = selected_owners
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("slide-removal owner count overflow".into()))?;
            } else {
                return refusal(
                    SlideRemovalRefusal::SharedOwner,
                    "another package part references the selected slide",
                );
            }
        }
    }
    if selected_owners != 1 {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected slide does not have exactly one presentation owner",
        );
    }
    Ok(())
}

/// The bound on relationships one removal proof scans.
pub(crate) fn relationship_scan_limit(limits: SourceLimits) -> Result<usize> {
    limits
        .max_parts()
        .checked_mul(16)
        .ok_or_else(|| Error::Invalid("slide-removal relationship limit overflow".into()))
}

fn checked_relationship_count(
    current: usize,
    limit: usize,
    package: &SourceBackedPackage,
) -> Result<usize> {
    package.check_execution()?;
    let next = current
        .checked_add(1)
        .ok_or_else(|| Error::Invalid("slide-removal relationship count overflow".into()))?;
    if next > limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Relationships),
            observed: next,
            maximum: limit,
        });
    }
    Ok(next)
}

fn reject_package_features(package: &SourceBackedPackage, presentation_xml: &[u8]) -> Result<()> {
    if package.has_encrypted_entries() {
        return refusal(
            SlideRemovalRefusal::UnknownPhysicalMember,
            "the source contains encrypted ZIP members",
        );
    }
    if !package.non_part_members().is_empty() {
        return refusal(
            SlideRemovalRefusal::UnknownPhysicalMember,
            "the source contains unknown non-Part members",
        );
    }
    if package.has_signature_infrastructure() {
        return refusal(
            SlideRemovalRefusal::SignedPackage,
            "digital-signature infrastructure requires an explicit signature policy",
        );
    }
    if has_macro_infrastructure(package) {
        return refusal(
            SlideRemovalRefusal::MacroEnabledPackage,
            "macro infrastructure is outside dependency-free slide removal",
        );
    }
    let text = std::str::from_utf8(presentation_xml)
        .map_err(|error| Error::Xml(format!("presentation XML is not UTF-8: {error}")))?;
    if crate::presentation_properties::metadata::protection::Settings::parse_xml(
        text,
        crate::presentation_properties::metadata::protection::Limits::from_source_limits(
            package.read_limits(),
        ),
    )?
    .is_protected()
    {
        return refusal(
            SlideRemovalRefusal::ProtectedPresentation,
            "the presentation has an active modify-password verifier",
        );
    }
    Ok(())
}

fn has_macro_infrastructure(package: &SourceBackedPackage) -> bool {
    package
        .rels()
        .iter()
        .any(|relationship| is_macro_relationship(relationship.reltype()))
        || package.iter_parts().any(|part| {
            matches!(
                part.content_type(),
                ct::PML_PRES_MACRO_MAIN
                    | ct::PML_SLIDESHOW_MACRO_MAIN
                    | ct::PML_TEMPLATE_MACRO_MAIN
                    | ct::OFC_VBA_PROJECT
                    | ct::OFC_VBA_PROJECT_SIGNATURE
                    | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
            ) || part
                .rels()
                .iter()
                .any(|relationship| is_macro_relationship(relationship.reltype()))
        })
}

fn is_macro_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
    )
}

fn reject_mce(xml: &[u8], package: &SourceBackedPackage) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        package.check_execution()?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MCE_NAMESPACE)
                {
                    return refusal(
                        SlideRemovalRefusal::MarkupCompatibility,
                        "the presentation contains markup-compatibility markup",
                    );
                }
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    // A declaration alone (LibreOffice declares the
                    // namespace on every root) carries no alternate content;
                    // MCE elements and attributes are refused where used.
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let attribute_namespace = reader.resolver().resolve_attribute(attribute.key).0;
                    if matches!(
                        attribute_namespace,
                        ResolveResult::Bound(Namespace(value)) if value == MCE_NAMESPACE
                    ) {
                        return refusal(
                            SlideRemovalRefusal::MarkupCompatibility,
                            "the presentation contains a markup-compatibility attribute",
                        );
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideRemovalRefusal::UnknownSemanticSurface,
                    "the presentation contains a document type or processing instruction",
                );
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

fn map_copy_refusal(value: Result<()>) -> Result<()> {
    match value {
        Err(Error::Refused(refusal)) if refusal.operation() == "slide_copy" => Err(Error::refused(
            refusal.kind(),
            "slide_removal",
            refusal.reason(),
        )),
        other => other,
    }
}

fn map_order_refusal(value: Result<()>) -> Result<()> {
    match value {
        Err(Error::Refused(refusal)) if refusal.operation() == "slide_order" => Err(
            Error::refused(refusal.kind(), "slide_removal", refusal.reason()),
        ),
        other => other,
    }
}

fn refusal<T>(kind: SlideRemovalRefusal, _detail: impl Into<String>) -> Result<T> {
    Err(Error::slide_removal_refused(kind))
}
