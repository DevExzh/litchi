//! Bounded planning for atomic dependency-free whole-slide removal.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::copy_plan::{
    has_signature_infrastructure, reject_mce, reject_unknown_non_part_members, resolve_slide,
    validate_registered_layout, validate_slide_surface,
};
use super::model::{Slide, Snapshot, invalid};
use super::patch::Patch;
use super::patch::SinglePartChange;
use crate::{Error, Result, SlideCopyRefusal, SlideRemovalRefusal};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const PATCH_MAGIC: &[u8; 8] = b"LPRM0001";
const PATCH_HEADER_BYTES: usize = PATCH_MAGIC.len() + 32 + 32 + 8;

/// Durable exact-source slide-removal patch bound to complete package revisions.
///
/// Unlike a general resource patch, this wrapper preserves the incoming-owner
/// proof across serialization: both forward and inverse publication require the
/// exact complete-package revision captured during planning.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideRemovalPatch {
    source_revision: [u8; 32],
    target_revision: [u8; 32],
    patch: Patch,
}

impl SlideRemovalPatch {
    /// Complete-package revision required before publication.
    #[must_use]
    pub const fn source_revision(&self) -> [u8; 32] {
        self.source_revision
    }

    /// Complete-package revision guaranteed after successful publication.
    #[must_use]
    pub const fn target_revision(&self) -> [u8; 32] {
        self.target_revision
    }

    /// Number of exact part resources in the write set.
    #[must_use]
    pub fn resource_count(&self) -> usize {
        self.patch.resource_count()
    }

    /// Physical part names in deterministic order.
    pub fn resources(&self) -> impl ExactSizeIterator<Item = &PackURI> {
        self.patch.resources()
    }

    /// Exact complete-revision-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_revision: self.target_revision,
            target_revision: self.source_revision,
            patch: self.patch.inverse(),
        }
    }

    /// Serialize this removal patch into the stable `LPRM0001` format.
    ///
    /// # Errors
    ///
    /// Returns an error for allocation, length overflow, or the inherited
    /// finite durable-patch bound.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let payload = self.patch.to_bytes()?;
        let length = PATCH_HEADER_BYTES
            .checked_add(payload.len())
            .ok_or_else(|| invalid("slide-removal durable patch length overflow"))?;
        let limit = self
            .patch
            .limits()
            .max_patch_bytes()
            .checked_add(PATCH_HEADER_BYTES)
            .ok_or_else(|| invalid("slide-removal durable patch limit overflow"))?;
        if length > limit {
            return Err(Error::Limit {
                resource: "slide-removal durable patch bytes",
                limit,
            });
        }
        let payload_len = u64::try_from(payload.len())
            .map_err(|_error| invalid("slide-removal durable patch length exceeds u64"))?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "slide-removal durable patch",
                source,
            })?;
        output.extend_from_slice(PATCH_MAGIC);
        output.extend_from_slice(&self.source_revision);
        output.extend_from_slice(&self.target_revision);
        output.extend_from_slice(&payload_len.to_le_bytes());
        output.extend_from_slice(&payload);
        Ok(output)
    }

    /// Parse a durable removal patch under conservative finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, trailing, or unbounded input.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, super::Limits::default())
    }

    /// Parse a durable removal patch under caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, trailing, or unbounded input.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: super::Limits) -> Result<Self> {
        let limit = limits
            .max_patch_bytes()
            .checked_add(PATCH_HEADER_BYTES)
            .ok_or_else(|| invalid("slide-removal durable patch limit overflow"))?;
        if bytes.len() > limit {
            return Err(Error::Limit {
                resource: "slide-removal durable patch bytes",
                limit,
            });
        }
        if bytes.len() < PATCH_HEADER_BYTES || &bytes[..PATCH_MAGIC.len()] != PATCH_MAGIC {
            return Err(invalid(
                "slide-removal durable patch has an unsupported version",
            ));
        }
        let mut position = PATCH_MAGIC.len();
        let source_revision = take_revision(bytes, &mut position)?;
        let target_revision = take_revision(bytes, &mut position)?;
        let length_bytes: [u8; 8] = bytes
            .get(position..position + 8)
            .ok_or_else(|| invalid("slide-removal durable patch is truncated"))?
            .try_into()
            .map_err(|_error| invalid("slide-removal durable patch length is malformed"))?;
        position += 8;
        let payload_len = usize::try_from(u64::from_le_bytes(length_bytes))
            .map_err(|_error| invalid("slide-removal durable patch length exceeds usize"))?;
        let end = position
            .checked_add(payload_len)
            .ok_or_else(|| invalid("slide-removal durable patch length overflow"))?;
        if end != bytes.len() {
            return Err(invalid(
                "slide-removal durable patch is truncated or has trailing bytes",
            ));
        }
        let patch = Patch::from_bytes_with_limits(&bytes[position..end], limits)?;
        if patch.exact_slide_removal_change().is_none() {
            return Err(invalid(
                "slide-removal durable patch is not an exact presentation-and-slide change",
            ));
        }
        Ok(Self {
            source_revision,
            target_revision,
            patch,
        })
    }
}

/// Immutable source-bound plan for removing one dependency-free slide.
///
/// The supported deletion boundary is deliberately narrow: the selected slide
/// may have exactly one internal relationship to a registered shared layout and
/// no other outgoing relationship. No other package resource may target the
/// slide. The durable patch therefore removes exactly the slide part and edits
/// exactly the presentation owner; every other resource remains outside the
/// write set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideRemovalPlan {
    source: Slide,
    position: usize,
    layout: PackURI,
    planned_bytes: usize,
    source_revision: [u8; 32],
    patch: SlideRemovalPatch,
}

impl SlideRemovalPlan {
    /// Source semantic slide identity captured by the immutable snapshot.
    #[must_use]
    pub const fn source(&self) -> &Slide {
        &self.source
    }

    /// Zero-based source position of the removed slide.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Existing shared layout retained without rewriting.
    #[must_use]
    pub const fn retained_layout(&self) -> &PackURI {
        &self.layout
    }

    /// Bounded bytes in the two-resource removal write set before encoding.
    #[must_use]
    pub const fn planned_bytes(&self) -> usize {
        self.planned_bytes
    }

    /// Fingerprint of the complete package graph against which this plan was built.
    #[must_use]
    pub const fn source_revision(&self) -> [u8; 32] {
        self.source_revision
    }

    /// Exact durable complete-revision-bound forward patch.
    ///
    /// The patch can be serialized with [`SlideRemovalPatch::to_bytes`] and
    /// restored after application with [`SlideRemovalPatch::inverse`].
    #[must_use]
    pub const fn patch(&self) -> &SlideRemovalPatch {
        &self.patch
    }
}

impl Snapshot {
    /// Plan removal of one dependency-free slide without mutation.
    ///
    /// Signed, macro-enabled, protected, MCE-bearing, and final-slide packages
    /// are refused. The selected slide must have exactly one registered layout
    /// relationship, no notes, charts, media, custom XML, external links, or
    /// other dependencies, and no incoming owner except its one presentation
    /// relationship.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an unsupported deletion boundary, or a
    /// limit/allocation/OPC/XML error while constructing the exact candidate.
    pub fn plan_slide_removal<'s>(
        &self,
        slide: impl Into<crate::slide::Key<'s>>,
    ) -> Result<SlideRemovalPlan> {
        let selected = resolve_slide(self, slide.into())?;
        let position = self
            .slides
            .iter()
            .position(|candidate| candidate.id == selected.id)
            .ok_or_else(|| invalid("slide-removal selection lost its source position"))?;
        if self.slides.len() == 1 {
            return refusal(
                SlideRemovalRefusal::FinalSlide,
                "a presentation must retain at least one slide",
            );
        }
        if self.package.part_count() > self.limits.max_parts() {
            return Err(Error::Limit {
                resource: "slide-removal package parts",
                limit: self.limits.max_parts(),
            });
        }
        map_copy_refusal(reject_unknown_non_part_members(
            self.package.as_ref(),
            "slide-removal source",
        ))?;
        if has_signature_infrastructure(self.package.as_ref()) {
            return refusal(
                SlideRemovalRefusal::SignedPackage,
                "digital-signature infrastructure requires an explicit signature policy",
            );
        }
        if has_macro_infrastructure(self.package.as_ref()) {
            return refusal(
                SlideRemovalRefusal::MacroEnabledPackage,
                "macro infrastructure is outside dependency-free slide removal",
            );
        }

        let presentation = self.package.get_part(&self.presentation_name)?;
        crate::parts::validate_content_type(presentation, ct::PML_PRESENTATION_MAIN)?;
        map_copy_refusal(reject_mce(presentation.blob(), "presentation owner"))?;
        let presentation_text = std::str::from_utf8(presentation.blob())
            .map_err(|error| Error::Xml(format!("presentation XML is not UTF-8: {error}")))?;
        if crate::presentation_properties::metadata::protection::Settings::parse_xml(
            presentation_text,
        )?
        .is_protected()
        {
            return refusal(
                SlideRemovalRefusal::ProtectedPresentation,
                "the presentation has an active modify-password verifier",
            );
        }
        validate_presentation_owner(presentation.blob(), &selected)?;

        let owner = self.package.get_part(&selected.part_name)?;
        crate::parts::validate_content_type(owner, ct::PML_SLIDE)?;
        map_copy_refusal(validate_slide_surface(owner.blob()))?;
        let layout = dependency_free_layout(self.package.as_ref(), owner)?;
        crate::master_layout::validate_master_layout_graph(self.package.as_ref())?;
        map_copy_refusal(validate_registered_layout(
            self.package.as_ref(),
            presentation,
            &layout,
        ))?;
        let max_relationships = self
            .limits
            .max_parts()
            .checked_mul(16)
            .ok_or_else(|| invalid("slide-removal relationship limit overflow"))?;
        validate_unique_incoming_owner(
            self.package.as_ref(),
            &self.presentation_name,
            &selected,
            max_relationships,
        )?;

        let planned_bytes = presentation
            .blob()
            .len()
            .checked_add(owner.blob().len())
            .and_then(|bytes| bytes.checked_add(self.presentation_name.as_str().len()))
            .and_then(|bytes| bytes.checked_add(selected.part_name.as_str().len()))
            .ok_or_else(|| invalid("slide-removal planned byte count overflow"))?;
        if planned_bytes > self.limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "slide-removal planned bytes",
                limit: self.limits.max_patch_bytes(),
            });
        }

        let candidate = build_candidate(self, &selected)?;
        let exact_patch = Patch::capture(
            self.package.as_ref(),
            &candidate,
            self.presentation_name.clone(),
            self.limits,
        )?;
        let mut resources: Vec<_> = exact_patch.resources().collect();
        resources.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        let mut expected = [&self.presentation_name, &selected.part_name];
        expected.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        if resources != expected || !exact_patch.removes_resource(&selected.part_name) {
            return Err(invalid(
                "slide-removal candidate escaped its exact two-resource write set",
            ));
        }
        let target_revision = super::model::package_fingerprint(&candidate)?;
        let patch = SlideRemovalPatch {
            source_revision: self.revision,
            target_revision,
            patch: exact_patch,
        };
        Ok(SlideRemovalPlan {
            source: selected,
            position,
            layout,
            planned_bytes,
            source_revision: self.revision,
            patch,
        })
    }
}

pub(crate) fn apply_patch(
    package: &mut OpcPackage,
    patch: &SlideRemovalPatch,
    operation: &'static str,
    physical_source_provenance: bool,
) -> Result<Snapshot> {
    if super::model::package_fingerprint(package)? != patch.source_revision {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "the complete package graph differs from the slide-removal patch source",
        });
    }
    let (direction, slide_name) =
        patch
            .patch
            .exact_slide_removal_change()
            .ok_or(Error::UnsafeEdit {
                operation,
                reason: "the durable patch is not an exact presentation-and-slide change",
            })?;
    let slide_name = slide_name.clone();
    match direction {
        SinglePartChange::Remove => {
            let source =
                super::model::capture(package, patch.patch.limits(), physical_source_provenance)?;
            let proof = replan_exact_removal(&source, &slide_name, operation)?;
            if !patch.has_same_semantics(proof.patch()) {
                return Err(Error::UnsafeEdit {
                    operation,
                    reason: "the durable patch does not match a freshly proven slide removal",
                });
            }
            super::patch::apply_exact_revision(
                package,
                &patch.patch,
                patch.target_revision,
                physical_source_provenance,
            )
        },
        SinglePartChange::Add => {
            let mut candidate = package.clone();
            let restored = super::patch::apply_exact_revision(
                &mut candidate,
                &patch.patch,
                patch.target_revision,
                physical_source_provenance,
            )?;
            let proof = replan_exact_removal(&restored, &slide_name, operation)?;
            if !patch.inverse().has_same_semantics(proof.patch()) {
                return Err(Error::UnsafeEdit {
                    operation,
                    reason: "the durable inverse does not restore a freshly removable slide",
                });
            }
            *package = candidate;
            Ok(restored)
        },
    }
}

impl SlideRemovalPatch {
    fn has_same_semantics(&self, other: &Self) -> bool {
        self.source_revision == other.source_revision
            && self.target_revision == other.target_revision
            && self.patch.has_same_changes(&other.patch)
    }
}

fn replan_exact_removal(
    snapshot: &Snapshot,
    slide_name: &PackURI,
    operation: &'static str,
) -> Result<SlideRemovalPlan> {
    let position = snapshot
        .slides()
        .iter()
        .position(|slide| slide.part_name() == slide_name)
        .ok_or(Error::UnsafeEdit {
            operation,
            reason: "the durable patch slide is not owned by the presentation",
        })?;
    snapshot.plan_slide_removal(position)
}

fn take_revision(bytes: &[u8], position: &mut usize) -> Result<[u8; 32]> {
    let end = position
        .checked_add(32)
        .ok_or_else(|| invalid("slide-removal durable patch position overflow"))?;
    let revision = bytes
        .get(*position..end)
        .ok_or_else(|| invalid("slide-removal durable patch is truncated"))?
        .try_into()
        .map_err(|_error| invalid("slide-removal durable patch revision is malformed"))?;
    *position = end;
    Ok(revision)
}

fn build_candidate(snapshot: &Snapshot, selected: &Slide) -> Result<OpcPackage> {
    let mut candidate = snapshot.package.as_ref().clone();
    let xml = super::xml::remove_slide(
        candidate.get_part(&snapshot.presentation_name)?.blob(),
        &snapshot.slides,
        selected.id,
    )?;
    {
        let presentation = candidate.get_part_mut(&snapshot.presentation_name)?;
        if presentation
            .rels_mut()
            .remove(&selected.relationship_id)
            .is_none()
        {
            return Err(invalid(
                "slide-removal presentation relationship disappeared",
            ));
        }
        presentation.set_blob(xml);
    }
    if !candidate.remove_part(&selected.part_name) {
        return Err(invalid("slide-removal selected part disappeared"));
    }
    let captured = super::model::capture(
        &candidate,
        snapshot.limits,
        snapshot.physical_source_provenance,
    )?;
    if captured.slides.len() + 1 != snapshot.slides.len()
        || captured
            .slides
            .iter()
            .zip(
                snapshot
                    .slides
                    .iter()
                    .filter(|slide| slide.id != selected.id),
            )
            .any(|(actual, expected)| actual != expected)
    {
        return Err(invalid(
            "slide-removal candidate changed retained slide identity or order",
        ));
    }
    Ok(candidate)
}

fn dependency_free_layout(package: &OpcPackage, slide: &dyn Part) -> Result<PackURI> {
    if slide.rels().len() != 1 {
        return refusal(
            SlideRemovalRefusal::UnsupportedRelationship,
            "the selected slide must have only one layout relationship",
        );
    }
    let relationship = slide
        .rels()
        .iter()
        .next()
        .ok_or_else(|| invalid("slide-removal relationship count changed"))?;
    if relationship.is_external()
        || !crate::parts::is_relationship_type(
            relationship.reltype(),
            rt::SLIDE_LAYOUT,
            "slideLayout",
        )
    {
        return refusal(
            SlideRemovalRefusal::UnsupportedRelationship,
            "the selected slide has a non-layout or external relationship",
        );
    }
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
    crate::parts::validate_content_type(package.get_part(&layout)?, ct::PML_SLIDE_LAYOUT)?;
    Ok(layout)
}

fn validate_unique_incoming_owner(
    package: &OpcPackage,
    presentation_name: &PackURI,
    selected: &Slide,
    max_relationships: usize,
) -> Result<()> {
    let mut scanned = 0_usize;
    let mut selected_owners = 0_usize;
    for relationship in package.rels().iter() {
        scanned = checked_relationship_count(scanned, max_relationships)?;
        if relationship.is_external() {
            continue;
        }
        if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
            return refusal(
                SlideRemovalRefusal::AmbiguousTopology,
                "an internal package-root relationship has a query or fragment",
            );
        }
        if relationship.target_partname()? == selected.part_name {
            return refusal(
                SlideRemovalRefusal::SharedOwner,
                "the package root directly references the selected slide",
            );
        }
    }
    for owner in package.iter_parts() {
        for relationship in owner.rels().iter() {
            scanned = checked_relationship_count(scanned, max_relationships)?;
            if relationship.is_external() {
                continue;
            }
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return refusal(
                    SlideRemovalRefusal::AmbiguousTopology,
                    "an internal package relationship has a query or fragment",
                );
            }
            if relationship.target_partname()? != selected.part_name {
                continue;
            }
            if owner.partname() == presentation_name
                && relationship.r_id() == selected.relationship_id
                && crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
            {
                selected_owners = selected_owners
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-removal owner count overflow"))?;
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

fn checked_relationship_count(current: usize, limit: usize) -> Result<usize> {
    let next = current
        .checked_add(1)
        .ok_or_else(|| invalid("slide-removal relationship count overflow"))?;
    if next > limit {
        return Err(Error::Limit {
            resource: "slide-removal relationship scan",
            limit,
        });
    }
    Ok(next)
}

fn validate_presentation_owner(xml: &[u8], selected: &Slide) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut selected_bindings = 0_usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let drawingml = match namespace {
                    ResolveResult::Bound(Namespace(value))
                        if value == PML || value == STRICT_PML =>
                    {
                        false
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == DML || value == STRICT_DML =>
                    {
                        true
                    },
                    ResolveResult::Bound(_) => {
                        return refusal(
                            SlideRemovalRefusal::UnknownSemanticSurface,
                            "the presentation contains an extension element namespace",
                        );
                    },
                    ResolveResult::Unknown(_) | ResolveResult::Unbound => {
                        return refusal(
                            SlideRemovalRefusal::UnknownSemanticSurface,
                            "the presentation contains an unbound element namespace",
                        );
                    },
                };
                let element_local = element.local_name();
                if drawingml && !known_drawingml_element(element_local.as_ref()) {
                    return refusal(
                        SlideRemovalRefusal::UnknownSemanticSurface,
                        "the presentation contains an unmodeled DrawingML element",
                    );
                }
                if !drawingml && !known_presentation_element(element_local.as_ref()) {
                    return refusal(
                        SlideRemovalRefusal::UnknownSemanticSurface,
                        "the presentation contains an unmodeled PresentationML element",
                    );
                }
                if !drawingml && element_local.as_ref() == b"sldRg" {
                    return refusal(
                        SlideRemovalRefusal::SharedOwner,
                        "a slideshow range refers to slides by numeric position",
                    );
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let (attribute_namespace, _) =
                        reader.resolver().resolve_attribute(attribute.key);
                    match attribute_namespace {
                        ResolveResult::Bound(Namespace(value))
                            if value == REL || value == STRICT_REL =>
                        {
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Xml(error.to_string()))?;
                            if value.as_ref() == selected.relationship_id {
                                if element.local_name().as_ref() != b"sldId" {
                                    return refusal(
                                        SlideRemovalRefusal::SharedOwner,
                                        "a presentation feature also references the selected slide",
                                    );
                                }
                                selected_bindings =
                                    selected_bindings.checked_add(1).ok_or_else(|| {
                                        invalid("slide-removal XML binding count overflow")
                                    })?;
                            }
                        },
                        ResolveResult::Bound(Namespace(value)) if value == XML => {},
                        ResolveResult::Unbound => {
                            if !known_unqualified_presentation_attribute(
                                drawingml,
                                element_local.as_ref(),
                                key,
                            ) {
                                return refusal(
                                    SlideRemovalRefusal::UnknownSemanticSurface,
                                    "the presentation contains an unmodeled unqualified attribute",
                                );
                            }
                        },
                        ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
                            return refusal(
                                SlideRemovalRefusal::UnknownSemanticSurface,
                                "the presentation contains an unmodeled attribute namespace",
                            );
                        },
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideRemovalRefusal::UnknownSemanticSurface,
                    "the presentation contains a document type or processing instruction",
                );
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if selected_bindings != 1 {
        return refusal(
            SlideRemovalRefusal::AmbiguousTopology,
            "the selected presentation relationship is not bound exactly once",
        );
    }
    Ok(())
}

fn known_drawingml_element(element: &[u8]) -> bool {
    matches!(
        element,
        b"defPPr"
            | b"defRPr"
            | b"lvl1pPr"
            | b"lvl2pPr"
            | b"lvl3pPr"
            | b"lvl4pPr"
            | b"lvl5pPr"
            | b"lvl6pPr"
            | b"lvl7pPr"
            | b"lvl8pPr"
            | b"lvl9pPr"
            | b"solidFill"
            | b"schemeClr"
            | b"latin"
            | b"ea"
            | b"cs"
    )
}

fn known_presentation_element(element: &[u8]) -> bool {
    matches!(
        element,
        b"presentation"
            | b"sldMasterIdLst"
            | b"sldMasterId"
            | b"notesMasterIdLst"
            | b"notesMasterId"
            | b"handoutMasterIdLst"
            | b"handoutMasterId"
            | b"sldIdLst"
            | b"sldId"
            | b"sldSz"
            | b"notesSz"
            | b"defaultTextStyle"
            | b"custShowLst"
            | b"custShow"
            | b"sldLst"
            | b"sld"
            | b"showPr"
            | b"sldRg"
    )
}

fn known_unqualified_presentation_attribute(drawingml: bool, element: &[u8], value: &[u8]) -> bool {
    if drawingml {
        return match element {
            b"defRPr" => matches!(value, b"lang" | b"sz" | b"kern"),
            b"lvl1pPr" | b"lvl2pPr" | b"lvl3pPr" | b"lvl4pPr" | b"lvl5pPr" | b"lvl6pPr"
            | b"lvl7pPr" | b"lvl8pPr" | b"lvl9pPr" => matches!(
                value,
                b"marL"
                    | b"algn"
                    | b"defTabSz"
                    | b"rtl"
                    | b"eaLnBrk"
                    | b"latinLnBrk"
                    | b"hangingPunct"
            ),
            b"schemeClr" => value == b"val",
            b"latin" | b"ea" | b"cs" => value == b"typeface",
            b"defPPr" | b"solidFill" => false,
            _ => false,
        };
    }
    match element {
        b"presentation" => matches!(
            value,
            b"saveSubsetFonts"
                | b"autoCompressPictures"
                | b"bookmarkIdSeed"
                | b"firstSlideNum"
                | b"showSpecialPlsOnTitleSld"
                | b"rtl"
                | b"removePersonalInfoOnSave"
                | b"compatMode"
                | b"strictFirstAndLastChars"
                | b"embedTrueTypeFonts"
        ),
        b"sldId" => value == b"id",
        b"sldMasterId" => value == b"id",
        b"sldSz" => matches!(value, b"cx" | b"cy" | b"type"),
        b"notesSz" => matches!(value, b"cx" | b"cy"),
        b"custShow" => matches!(value, b"name" | b"id"),
        // The dependency-free removal planner intentionally supports only the
        // finite direct presentation vocabulary emitted by this crate. Other
        // PML attributes require a typed owner before slide deletion is safe.
        _ => false,
    }
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
            ct::PML_PRES_MACRO_MAIN
                | ct::PML_SLIDESHOW_MACRO_MAIN
                | ct::PML_TEMPLATE_MACRO_MAIN
                | ct::OFC_VBA_PROJECT
                | ct::OFC_VBA_PROJECT_SIGNATURE
                | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
        ) || part.rels().iter().any(|relationship| {
            matches!(
                relationship.reltype(),
                rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
            )
        })
    })
}

fn map_copy_refusal(value: Result<()>) -> Result<()> {
    match value {
        Ok(()) => Ok(()),
        Err(Error::SlideCopyPlan { kind, detail }) => {
            let kind = match kind {
                SlideCopyRefusal::SignedPackage => SlideRemovalRefusal::SignedPackage,
                SlideCopyRefusal::ProtectedPresentation => {
                    SlideRemovalRefusal::ProtectedPresentation
                },
                SlideCopyRefusal::MarkupCompatibility => SlideRemovalRefusal::MarkupCompatibility,
                SlideCopyRefusal::UnknownSemanticSurface => {
                    SlideRemovalRefusal::UnknownSemanticSurface
                },
                SlideCopyRefusal::SharedOwner | SlideCopyRefusal::GlobalTableStyle => {
                    SlideRemovalRefusal::SharedOwner
                },
                SlideCopyRefusal::UnknownPhysicalMember => {
                    SlideRemovalRefusal::UnknownPhysicalMember
                },
                SlideCopyRefusal::UnsupportedRelationship => {
                    SlideRemovalRefusal::UnsupportedRelationship
                },
                SlideCopyRefusal::DependencyCycle | SlideCopyRefusal::AmbiguousTopology => {
                    SlideRemovalRefusal::AmbiguousTopology
                },
            };
            refusal(kind, detail)
        },
        Err(error) => Err(error),
    }
}

fn refusal<T>(kind: SlideRemovalRefusal, detail: impl Into<String>) -> Result<T> {
    Err(Error::SlideRemovalPlan {
        kind,
        detail: detail.into(),
    })
}
