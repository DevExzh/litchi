//! Bounded planning for an atomic same-package whole-slide copy operation.
//!
//! A slide is not an isolated XML part.  Its private charts, media, embedded
//! packages, and diagram resources can be copied, while its layout is a shared
//! presentation owner and notes/comments can depend on catalogs outside the
//! slide's outgoing OPC graph. This module inventories only the closure that
//! the current package model can prove independent, then builds the complete
//! exact-resource candidate without publishing it.

use std::collections::{HashMap, HashSet, VecDeque};
use std::fmt::Write as _;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Slide, Snapshot, invalid};
use super::patch::Patch;
use crate::{Error, Result, SlideCopyRefusal};

const MIN_SLIDE_ID: u32 = 256;
const MAX_SLIDE_ID: u32 = 2_147_483_647;
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const CHART: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const DIAGRAM: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const STRICT_DIAGRAM: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const CHART_DRAWING: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
const STRICT_CHART_DRAWING: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chartDrawing";
const DIAGRAM_DRAWING: &[u8] = b"http://schemas.microsoft.com/office/drawing/2008/diagram";
const CHART_STYLE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const CHART_COLOR: &[u8] = b"http://schemas.microsoft.com/office/drawing/2012/chartColorStyle";
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";

/// One raw part in a dependency-closed whole-slide copy plan.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideCopyPart {
    source: PackURI,
    target: PackURI,
    content_type: String,
    bytes: usize,
    relationships: usize,
}

impl SlideCopyPart {
    /// Existing source part whose exact bytes and relationship IDs would be retained.
    #[must_use]
    pub const fn source(&self) -> &PackURI {
        &self.source
    }

    /// Deterministic collision-free part name reserved by this plan.
    #[must_use]
    pub const fn target(&self) -> &PackURI {
        &self.target
    }

    /// Exact source content type to retain.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact source payload size.
    #[must_use]
    pub const fn bytes(&self) -> usize {
        self.bytes
    }

    /// Number of raw outgoing relationships retained with this part.
    #[must_use]
    pub const fn relationship_count(&self) -> usize {
        self.relationships
    }
}

/// Immutable source-bound plan for the largest proven whole-slide closure.
///
/// Planning is intentionally non-publishing. It proves a distinct slide ID,
/// deterministic collision names, one reusable layout boundary, an acyclic
/// private dependency graph, inert external relationships, and finite resource
/// use.  Arbitrary equal parts are never deduplicated: byte equality alone does
/// not prove equivalent ownership or relationship semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideCopyPlan {
    source: Slide,
    position: usize,
    slide_id: u32,
    presentation_relationship_id: String,
    parts: Box<[SlideCopyPart]>,
    reused_layout: PackURI,
    external_relationships: usize,
    planned_bytes: usize,
    source_revision: [u8; 32],
    patch: Patch,
}

impl SlideCopyPlan {
    /// Source semantic slide identity captured by the immutable snapshot.
    #[must_use]
    pub const fn source(&self) -> &Slide {
        &self.source
    }

    /// Checked zero-based insertion position in the source presentation.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Collision-free `p:sldId@id` reserved by this plan.
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Collision-free presentation relationship ID reserved by this plan.
    #[must_use]
    pub fn presentation_relationship_id(&self) -> &str {
        &self.presentation_relationship_id
    }

    /// Raw slide and private dependency parts in deterministic source-name order.
    #[must_use]
    pub fn parts(&self) -> &[SlideCopyPart] {
        &self.parts
    }

    /// Existing layout that a future copy must reuse without rewriting it.
    #[must_use]
    pub const fn reused_layout(&self) -> &PackURI {
        &self.reused_layout
    }

    /// Number of external targets retained inertly and never dereferenced.
    #[must_use]
    pub const fn external_relationship_count(&self) -> usize {
        self.external_relationships
    }

    /// Bounded source payload and relationship-metadata bytes in the inventory.
    #[must_use]
    pub const fn planned_bytes(&self) -> usize {
        self.planned_bytes
    }

    /// Fingerprint of the complete package graph against which this plan was built.
    #[must_use]
    pub const fn source_revision(&self) -> [u8; 32] {
        self.source_revision
    }

    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }
}

impl Snapshot {
    /// Plan a dependency-closed copy of one existing slide without mutation.
    ///
    /// The returned plan contains a fully validated exact-resource candidate. The
    /// supported closure contains the raw slide plus acyclic image, audio,
    /// video, chart, chart-drawing, diagram, tag, theme-override, OLE, and
    /// embedded-package dependencies.  One existing layout is reused as a
    /// validated shared boundary.  Explicitly allowlisted external media,
    /// object, image, and hyperlink relationships are retained inertly.
    ///
    /// Notes, comments, cross-slide links, tables, controls, model3d/content
    /// parts, unknown internal relationship families, unknown slide extension
    /// namespaces, MCE, protection, signatures, and cycles are refused before
    /// any output exists.  Macros outside the selected closure remain inert and
    /// untouched; a macro/control edge from the slide is outside this closure.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an unsupported ownership boundary, or a
    /// limit/allocation/OPC/XML error while constructing the bounded inventory.
    pub fn plan_slide_copy<'s>(
        &self,
        source: impl Into<crate::slide::Key<'s>>,
        position: usize,
    ) -> Result<SlideCopyPlan> {
        if position > self.slides.len() {
            return Err(Error::SlideIndexOutOfBounds {
                index: position,
                len: self.slides.len().saturating_add(1),
            });
        }
        let source = resolve_slide(self, source.into())?;
        if has_signature_infrastructure(self.package.as_ref()) {
            return refusal(
                SlideCopyRefusal::SignedPackage,
                "digital-signature infrastructure must be handled by an explicit signature policy",
            );
        }
        let presentation = self.package.get_part(&self.presentation_name)?;
        reject_mce(presentation.blob(), "presentation owner")?;
        let presentation_text = std::str::from_utf8(presentation.blob())
            .map_err(|error| Error::Xml(format!("presentation XML is not UTF-8: {error}")))?;
        if crate::presentation_properties::metadata::protection::Settings::parse_xml(
            presentation_text,
        )?
        .is_protected()
        {
            return refusal(
                SlideCopyRefusal::ProtectedPresentation,
                "the presentation has an active modify-password verifier",
            );
        }
        validate_slide_ids(&self.slides)?;
        if self.slides.len() == crate::parts::MAX_SLIDES {
            return Err(Error::Limit {
                resource: "slide-copy presentation slides",
                limit: crate::parts::MAX_SLIDES,
            });
        }
        // This validates that the reused layout is registered through exactly
        // the package's modeled master/layout/theme owners.  No shared owner is
        // copied or rewritten by this plan.
        crate::master_layout::validate_master_layout_graph(self.package.as_ref())?;

        let slide = self.package.get_part(&source.part_name)?;
        crate::parts::validate_content_type(slide, ct::PML_SLIDE)?;
        validate_slide_surface(slide.blob())?;

        let (owned, edges, reused_layout, external_relationships, planned_bytes) =
            collect_owned_closure(self.package.as_ref(), &source.part_name, self.limits)?;
        reject_cycles(&owned, &edges)?;
        validate_registered_layout(self.package.as_ref(), presentation, &reused_layout)?;
        if self.package.part_count() > self.limits.max_parts() {
            return Err(Error::Limit {
                resource: "slide-copy package parts",
                limit: self.limits.max_parts(),
            });
        }

        let mut names: Vec<_> = owned.into_iter().collect();
        names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(names.len())
            .map_err(|source| Error::Allocation {
                resource: "slide-copy planned parts",
                source,
            })?;
        for name in names {
            let part = self.package.get_part(&name)?;
            parts.push(SlideCopyPart {
                target: available_copy_name(self.package.as_ref(), &name, self.limits.max_parts())?,
                source: name,
                content_type: try_copy_string(
                    part.content_type(),
                    "slide-copy planned content types",
                )?,
                bytes: part.blob().len(),
                relationships: part.rels().len(),
            });
        }
        let slide_id = next_slide_id(&self.slides)?;
        let presentation_relationship_id = next_relationship_id(presentation.rels())?;
        preflight_copy_candidate(self, &parts, presentation, planned_bytes)?;
        let candidate = build_copy_candidate(
            self,
            &source,
            position,
            slide_id,
            &presentation_relationship_id,
            &reused_layout,
            &parts,
        )?;
        let patch = Patch::capture(
            self.package.as_ref(),
            &candidate,
            self.presentation_name.clone(),
            self.limits,
        )?;
        Ok(SlideCopyPlan {
            source,
            position,
            slide_id,
            presentation_relationship_id,
            parts: parts.into_boxed_slice(),
            reused_layout,
            external_relationships,
            planned_bytes,
            source_revision: self.revision,
            patch,
        })
    }
}

fn preflight_copy_candidate(
    snapshot: &Snapshot,
    parts: &[SlideCopyPart],
    presentation: &dyn Part,
    planned_bytes: usize,
) -> Result<()> {
    let resulting_parts = snapshot
        .package
        .part_count()
        .checked_add(parts.len())
        .ok_or_else(|| invalid("slide-copy resulting part count overflow"))?;
    if resulting_parts > snapshot.limits.max_parts() {
        return Err(Error::Limit {
            resource: "slide-copy resulting package parts",
            limit: snapshot.limits.max_parts(),
        });
    }
    let presentation_bytes = checked_plan_bytes(0, presentation)?;
    let conservative_patch_bytes = planned_bytes
        .checked_mul(2)
        .and_then(|value| {
            presentation_bytes
                .checked_mul(2)
                .and_then(|owner| value.checked_add(owner))
        })
        .and_then(|value| {
            parts.iter().try_fold(value, |total, part| {
                total
                    .checked_add(part.target.as_str().len())
                    .and_then(|next| next.checked_add(part.content_type.len()))
            })
        })
        .and_then(|value| {
            parts
                .len()
                .checked_mul(32)
                .and_then(|n| value.checked_add(n))
        })
        .and_then(|value| value.checked_add(128))
        .ok_or_else(|| invalid("slide-copy candidate byte count overflow"))?;
    if conservative_patch_bytes > snapshot.limits.max_patch_bytes() {
        return Err(Error::Limit {
            resource: "slide-copy candidate patch bytes",
            limit: snapshot.limits.max_patch_bytes(),
        });
    }
    let mut targets = HashSet::new();
    targets
        .try_reserve(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy target identities",
            source,
        })?;
    for part in parts {
        snapshot.package.validate_new_part_name(&part.target)?;
        if !targets.insert(part.target.clone()) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "two copied resources selected the same target part name",
            );
        }
    }
    Ok(())
}

fn build_copy_candidate(
    snapshot: &Snapshot,
    source_slide: &Slide,
    position: usize,
    slide_id: u32,
    presentation_relationship_id: &str,
    reused_layout: &PackURI,
    parts: &[SlideCopyPart],
) -> Result<OpcPackage> {
    let mut mapping = HashMap::new();
    mapping
        .try_reserve(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy part-name mapping",
            source,
        })?;
    for part in parts {
        if mapping
            .insert(part.source.clone(), part.target.clone())
            .is_some()
        {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "the copied closure repeats a source part",
            );
        }
    }
    let copied_slide = mapping
        .get(&source_slide.part_name)
        .cloned()
        .ok_or_else(|| invalid("slide-copy candidate omitted the selected slide"))?;

    // All topology, byte, part-count, name, and ID checks have completed before
    // this bounded graph clone. Blob allocations remain shared through Arc.
    let mut candidate = snapshot.package.as_ref().clone();
    for planned in parts {
        let source = snapshot.package.get_part(&planned.source)?;
        let mut copied = BlobPart::new_shared(
            planned.target.clone(),
            planned.content_type.clone(),
            source.blob_arc(),
        );
        for relationship in source.rels().iter() {
            let (target, mode) = if relationship.is_external() {
                (relationship.target_ref().to_owned(), TargetMode::External)
            } else {
                let source_target = relationship.target_partname()?;
                let copied_target = if source_target == *reused_layout
                    && planned.source == source_slide.part_name
                {
                    reused_layout
                } else {
                    mapping
                        .get(&source_target)
                        .ok_or_else(|| Error::SlideCopyPlan {
                            kind: SlideCopyRefusal::AmbiguousTopology,
                            detail: "an internal copied relationship escaped the planned closure"
                                .to_owned(),
                        })?
                };
                (
                    copied_target.relative_ref(planned.target.base_uri()),
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

    let presentation = snapshot.package.get_part(&snapshot.presentation_name)?;
    let source_relationship = presentation
        .rels()
        .get(&source_slide.relationship_id)
        .ok_or_else(|| invalid("slide-copy source presentation relationship disappeared"))?;
    let xml = super::xml::insert_slide(
        presentation.blob(),
        &snapshot.slides,
        position,
        slide_id,
        presentation_relationship_id,
    )?;
    {
        let staged = candidate.get_part_mut(&snapshot.presentation_name)?;
        staged.rels_mut().try_add_relationship(
            source_relationship.reltype().to_owned(),
            copied_slide.relative_ref(snapshot.presentation_name.base_uri()),
            presentation_relationship_id.to_owned(),
            TargetMode::Internal,
        )?;
        staged.set_blob(xml);
    }
    let captured = super::model::capture(&candidate, snapshot.limits)?;
    let published = captured
        .slides
        .get(position)
        .ok_or_else(|| invalid("slide-copy candidate lost its insertion position"))?;
    if published.id != slide_id || published.part_name != copied_slide {
        return Err(invalid(
            "slide-copy candidate did not publish the reserved slide identity",
        ));
    }
    Ok(candidate)
}

type Closure = (
    HashSet<PackURI>,
    HashMap<PackURI, Vec<PackURI>>,
    PackURI,
    usize,
    usize,
);

fn collect_owned_closure(
    package: &OpcPackage,
    slide: &PackURI,
    limits: super::Limits,
) -> Result<Closure> {
    let mut queue = VecDeque::new();
    queue.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "slide-copy dependency queue",
        source,
    })?;
    queue.push_back(slide.clone());
    let mut owned = HashSet::new();
    owned
        .try_reserve(limits.max_parts().min(package.part_count()))
        .map_err(|source| Error::Allocation {
            resource: "slide-copy dependency identities",
            source,
        })?;
    let mut edges = HashMap::new();
    edges
        .try_reserve(limits.max_parts().min(package.part_count()))
        .map_err(|source| Error::Allocation {
            resource: "slide-copy dependency edges",
            source,
        })?;
    let mut reused_layout = None;
    let mut external_relationships = 0usize;
    let mut planned_bytes = 0usize;

    while let Some(name) = queue.pop_front() {
        if !owned.insert(name.clone()) {
            continue;
        }
        if owned.len() > limits.max_parts() {
            return Err(Error::Limit {
                resource: "slide-copy dependency parts",
                limit: limits.max_parts(),
            });
        }
        let part = package.get_part(&name)?;
        planned_bytes = checked_plan_bytes(planned_bytes, part)?;
        if planned_bytes > limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "slide-copy planned closure bytes",
                limit: limits.max_patch_bytes(),
            });
        }
        let mut outgoing = Vec::new();
        outgoing
            .try_reserve_exact(part.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "slide-copy outgoing dependency edges",
                source,
            })?;
        for relationship in part.rels().iter() {
            if relationship.target_ref().is_empty() {
                return refusal(
                    SlideCopyRefusal::AmbiguousTopology,
                    "a dependency relationship has an empty target",
                );
            }
            let relationship_type = relationship.reltype();
            if crate::parts::is_relationship_type(
                relationship_type,
                rt::SLIDE_LAYOUT,
                "slideLayout",
            ) {
                if relationship.is_external() {
                    return refusal(
                        SlideCopyRefusal::SharedOwner,
                        "a slide-layout relationship cannot have an external target",
                    );
                }
                if name != *slide {
                    return refusal(
                        SlideCopyRefusal::SharedOwner,
                        "a private dependency reaches a shared slide layout",
                    );
                }
                let target = relationship.target_partname()?;
                if reused_layout.replace(target.clone()).is_some() {
                    return refusal(
                        SlideCopyRefusal::AmbiguousTopology,
                        "the selected slide has more than one layout relationship",
                    );
                }
                let layout = package.get_part(&target)?;
                crate::parts::validate_content_type(layout, ct::PML_SLIDE_LAYOUT)?;
                continue;
            }
            if is_shared_owner_relationship(relationship_type) {
                return refusal(
                    SlideCopyRefusal::SharedOwner,
                    "the selected closure reaches notes, comments, another slide, or a presentation owner",
                );
            }
            if relationship.is_external() {
                if !is_allowed_external_relationship(relationship_type) {
                    return refusal(
                        SlideCopyRefusal::UnsupportedRelationship,
                        "an external dependency relationship is outside the inert external allowlist",
                    );
                }
                external_relationships = external_relationships
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-copy external relationship count overflow"))?;
                continue;
            }
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return refusal(
                    SlideCopyRefusal::AmbiguousTopology,
                    "an internal dependency target contains a query or fragment",
                );
            }
            let target = relationship.target_partname()?;
            let Some(role) = owned_role(relationship_type) else {
                return refusal(
                    SlideCopyRefusal::UnsupportedRelationship,
                    "an internal dependency relationship is outside the supported copy closure",
                );
            };
            let target_part = package.get_part(&target)?;
            validate_owned_target(role, target_part)?;
            validate_owned_surface(role, target_part)?;
            outgoing.push(target.clone());
            queue.push_back(target);
        }
        edges.insert(name, outgoing);
    }
    let reused_layout = reused_layout.ok_or_else(|| Error::SlideCopyPlan {
        kind: SlideCopyRefusal::AmbiguousTopology,
        detail: "the selected slide does not have exactly one reusable layout".to_owned(),
    })?;
    Ok((
        owned,
        edges,
        reused_layout,
        external_relationships,
        planned_bytes,
    ))
}

fn checked_plan_bytes(total: usize, part: &dyn Part) -> Result<usize> {
    let mut next = total
        .checked_add(part.partname().as_str().len())
        .and_then(|value| value.checked_add(part.content_type().len()))
        .and_then(|value| value.checked_add(part.blob().len()))
        .ok_or_else(|| invalid("slide-copy planned byte count overflow"))?;
    for relationship in part.rels().iter() {
        next = next
            .checked_add(relationship.r_id().len())
            .and_then(|value| value.checked_add(relationship.reltype().len()))
            .and_then(|value| value.checked_add(relationship.target_ref().len()))
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| invalid("slide-copy relationship byte count overflow"))?;
    }
    Ok(next)
}

#[derive(Clone, Copy)]
enum OwnedRole {
    Image,
    Audio,
    Video,
    Chart,
    ChartDrawing,
    DiagramData,
    DiagramLayout,
    DiagramStyle,
    DiagramColor,
    DiagramDrawing,
    Tag,
    ThemeOverride,
    Ole,
    Package,
    ChartStyle,
    ChartColor,
}

fn owned_role(value: &str) -> Option<OwnedRole> {
    if office_relationship(value, "image") {
        Some(OwnedRole::Image)
    } else if office_relationship(value, "audio") {
        Some(OwnedRole::Audio)
    } else if office_relationship(value, "video") || value == rt::MEDIA {
        Some(OwnedRole::Video)
    } else if office_relationship(value, "chart") {
        Some(OwnedRole::Chart)
    } else if office_relationship(value, "chartUserShapes") || office_relationship(value, "drawing")
    {
        Some(OwnedRole::ChartDrawing)
    } else if office_relationship(value, "diagramData") {
        Some(OwnedRole::DiagramData)
    } else if office_relationship(value, "diagramLayout") {
        Some(OwnedRole::DiagramLayout)
    } else if office_relationship(value, "diagramQuickStyle") {
        Some(OwnedRole::DiagramStyle)
    } else if office_relationship(value, "diagramColors") {
        Some(OwnedRole::DiagramColor)
    } else if office_relationship(value, "diagramDrawing") {
        Some(OwnedRole::DiagramDrawing)
    } else if office_relationship(value, "tags") {
        Some(OwnedRole::Tag)
    } else if office_relationship(value, "themeOverride") {
        Some(OwnedRole::ThemeOverride)
    } else if office_relationship(value, "oleObject") {
        Some(OwnedRole::Ole)
    } else if office_relationship(value, "package") {
        Some(OwnedRole::Package)
    } else if value == crate::chart::style::STYLE_RELATIONSHIP_TYPE
        || office_relationship(value, "chartStyle")
    {
        Some(OwnedRole::ChartStyle)
    } else if value == crate::chart::style::COLOR_RELATIONSHIP_TYPE
        || office_relationship(value, "chartColorStyle")
    {
        Some(OwnedRole::ChartColor)
    } else {
        None
    }
}

fn validate_owned_target(role: OwnedRole, part: &dyn Part) -> Result<()> {
    let content_type = part.content_type();
    let valid = match role {
        OwnedRole::Image => content_type.starts_with("image/"),
        OwnedRole::Audio => content_type.starts_with("audio/"),
        OwnedRole::Video => {
            content_type.starts_with("video/") || content_type.starts_with("audio/")
        },
        OwnedRole::Chart => content_type == ct::DML_CHART,
        OwnedRole::ChartDrawing => matches!(
            content_type,
            ct::DML_CHARTSHAPES | ct::OFC_DRAWING | ct::DML_DIAGRAM_DRAWING
        ),
        OwnedRole::DiagramData => content_type == ct::DML_DIAGRAM_DATA,
        OwnedRole::DiagramLayout => content_type == ct::DML_DIAGRAM_LAYOUT,
        OwnedRole::DiagramStyle => content_type == ct::DML_DIAGRAM_STYLE,
        OwnedRole::DiagramColor => content_type == ct::DML_DIAGRAM_COLORS,
        OwnedRole::DiagramDrawing => content_type == ct::DML_DIAGRAM_DRAWING,
        OwnedRole::Tag => {
            content_type == "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
        },
        OwnedRole::ThemeOverride => content_type == ct::OFC_THEME_OVERRIDE,
        OwnedRole::Ole => content_type == ct::OFC_OLE_OBJECT,
        OwnedRole::Package => true,
        OwnedRole::ChartStyle => content_type == crate::chart::style::STYLE_CONTENT_TYPE,
        OwnedRole::ChartColor => content_type == crate::chart::style::COLOR_CONTENT_TYPE,
    };
    if valid {
        Ok(())
    } else {
        refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "a dependency relationship targets an incompatible content type",
        )
    }
}

fn validate_owned_surface(role: OwnedRole, part: &dyn Part) -> Result<()> {
    let surface = match role {
        OwnedRole::Chart => Some(XmlSurface::Chart),
        OwnedRole::ChartDrawing => Some(XmlSurface::ChartDrawing),
        OwnedRole::DiagramData => Some(XmlSurface::DiagramData),
        OwnedRole::DiagramLayout => Some(XmlSurface::DiagramLayout),
        OwnedRole::DiagramStyle => Some(XmlSurface::DiagramStyle),
        OwnedRole::DiagramColor => Some(XmlSurface::DiagramColor),
        OwnedRole::DiagramDrawing => Some(XmlSurface::DiagramDrawing),
        OwnedRole::Tag => Some(XmlSurface::Tag),
        OwnedRole::ThemeOverride => Some(XmlSurface::ThemeOverride),
        OwnedRole::ChartStyle => Some(XmlSurface::ChartStyle),
        OwnedRole::ChartColor => Some(XmlSurface::ChartColor),
        OwnedRole::Image
        | OwnedRole::Audio
        | OwnedRole::Video
        | OwnedRole::Ole
        | OwnedRole::Package => None,
    };
    if let Some(surface) = surface {
        validate_xml_surface(part.blob(), surface)?;
    } else if is_xml_content_type(part.content_type()) {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "an otherwise opaque dependency uses an unmodeled XML media or package format",
        );
    }
    Ok(())
}

fn is_shared_owner_relationship(value: &str) -> bool {
    office_relationship(value, "slide")
        || office_relationship(value, "notesSlide")
        || office_relationship(value, "notesMaster")
        || office_relationship(value, "comments")
        || office_relationship(value, "commentAuthors")
        || office_relationship(value, "slideMaster")
        || office_relationship(value, "theme")
        || office_relationship(value, "tableStyles")
        || value == crate::modern_comments::MODERN_COMMENT_RELATIONSHIP_TYPE
        || value == crate::modern_comments::MODERN_COMMENT_AUTHOR_RELATIONSHIP_TYPE
}

fn is_allowed_external_relationship(value: &str) -> bool {
    office_relationship(value, "hyperlink")
        || office_relationship(value, "image")
        || office_relationship(value, "audio")
        || office_relationship(value, "video")
        || office_relationship(value, "oleObject")
        || value == rt::MEDIA
}

fn office_relationship(value: &str, local: &str) -> bool {
    value
        .strip_prefix("http://schemas.openxmlformats.org/officeDocument/2006/relationships/")
        .is_some_and(|actual| actual == local)
        || value
            .strip_prefix("http://purl.oclc.org/ooxml/officeDocument/relationships/")
            .is_some_and(|actual| actual == local)
}

fn reject_cycles(owned: &HashSet<PackURI>, edges: &HashMap<PackURI, Vec<PackURI>>) -> Result<()> {
    let mut indegree = HashMap::new();
    indegree
        .try_reserve(owned.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy dependency indegrees",
            source,
        })?;
    for name in owned {
        indegree.insert(name.clone(), 0usize);
    }
    for targets in edges.values() {
        for target in targets {
            let count = indegree.get_mut(target).ok_or_else(|| {
                invalid("slide-copy dependency closure is missing an edge target")
            })?;
            *count = count
                .checked_add(1)
                .ok_or_else(|| invalid("slide-copy dependency indegree overflow"))?;
        }
    }
    let mut queue = VecDeque::new();
    queue
        .try_reserve(indegree.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy cycle queue",
            source,
        })?;
    for (name, count) in &indegree {
        if *count == 0 {
            queue.push_back(name.clone());
        }
    }
    let mut visited = 0usize;
    while let Some(name) = queue.pop_front() {
        visited = visited
            .checked_add(1)
            .ok_or_else(|| invalid("slide-copy cycle visit count overflow"))?;
        if let Some(targets) = edges.get(&name) {
            for target in targets {
                let count = indegree
                    .get_mut(target)
                    .ok_or_else(|| invalid("slide-copy dependency closure lost an edge target"))?;
                *count = count
                    .checked_sub(1)
                    .ok_or_else(|| invalid("slide-copy dependency indegree underflow"))?;
                if *count == 0 {
                    queue.push_back(target.clone());
                }
            }
        }
    }
    if visited == owned.len() {
        Ok(())
    } else {
        refusal(
            SlideCopyRefusal::DependencyCycle,
            "the private dependency graph contains a cycle",
        )
    }
}

fn validate_slide_surface(xml: &[u8]) -> Result<()> {
    validate_xml_surface(xml, XmlSurface::Slide)
}

#[derive(Clone, Copy)]
enum XmlSurface {
    Slide,
    Chart,
    ChartDrawing,
    DiagramData,
    DiagramLayout,
    DiagramStyle,
    DiagramColor,
    DiagramDrawing,
    Tag,
    ThemeOverride,
    ChartStyle,
    ChartColor,
}

impl XmlSurface {
    const fn label(self) -> &'static str {
        match self {
            Self::Slide => "selected slide",
            Self::Chart => "chart dependency",
            Self::ChartDrawing => "chart-drawing dependency",
            Self::DiagramData => "diagram-data dependency",
            Self::DiagramLayout => "diagram-layout dependency",
            Self::DiagramStyle => "diagram-style dependency",
            Self::DiagramColor => "diagram-color dependency",
            Self::DiagramDrawing => "diagram-drawing dependency",
            Self::Tag => "tag dependency",
            Self::ThemeOverride => "theme-override dependency",
            Self::ChartStyle => "chart-style dependency",
            Self::ChartColor => "chart-color dependency",
        }
    }

    const fn root(self) -> (&'static [u8], &'static [u8]) {
        match self {
            Self::Slide => (PML, b"sld"),
            Self::Chart => (CHART, b"chartSpace"),
            Self::ChartDrawing => (CHART_DRAWING, b"userShapes"),
            Self::DiagramData => (DIAGRAM, b"dataModel"),
            Self::DiagramLayout => (DIAGRAM, b"layoutDef"),
            Self::DiagramStyle => (DIAGRAM, b"styleDef"),
            Self::DiagramColor => (DIAGRAM, b"colorsDef"),
            Self::DiagramDrawing => (DIAGRAM_DRAWING, b"drawing"),
            Self::Tag => (PML, b"tagLst"),
            Self::ThemeOverride => (DML, b"themeOverride"),
            Self::ChartStyle => (CHART_STYLE, b"chartStyle"),
            Self::ChartColor => (CHART_COLOR, b"colorStyle"),
        }
    }
}

fn validate_xml_surface(xml: &[u8], surface: XmlSurface) -> Result<()> {
    reject_mce(xml, surface.label())?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut root_seen = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let namespace = match namespace {
                    ResolveResult::Bound(Namespace(value)) => value,
                    ResolveResult::Unknown(_) | ResolveResult::Unbound => {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "the selected slide contains an unbound element namespace",
                        );
                    },
                };
                if !root_seen {
                    let (expected_namespace, expected_local_name) = surface.root();
                    if !equivalent_surface_namespace(namespace, expected_namespace)
                        || element.local_name().as_ref() != expected_local_name
                    {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            format!("{} has an unexpected document element", surface.label()),
                        );
                    }
                    root_seen = true;
                }
                if !allowed_surface_element_namespace(surface, namespace) {
                    return refusal(
                        SlideCopyRefusal::UnknownSemanticSurface,
                        format!(
                            "{} contains an unmodeled extension namespace",
                            surface.label()
                        ),
                    );
                }
                if (namespace == DML || namespace == STRICT_DML)
                    && element.local_name().as_ref() == b"tbl"
                {
                    return refusal(
                        SlideCopyRefusal::GlobalTableStyle,
                        "DrawingML tables can reference presentation-global style identities without an OPC edge",
                    );
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }
                    match reader.resolver().resolve_attribute(attribute.key).0 {
                        ResolveResult::Unbound => {},
                        ResolveResult::Bound(Namespace(value))
                            if allowed_surface_attribute_namespace(surface, value) => {},
                        ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
                            return refusal(
                                SlideCopyRefusal::UnknownSemanticSurface,
                                format!(
                                    "{} contains an unmodeled qualified attribute",
                                    surface.label()
                                ),
                            );
                        },
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::UnknownSemanticSurface,
                    format!(
                        "{} contains a document type or processing instruction",
                        surface.label()
                    ),
                );
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if root_seen {
        Ok(())
    } else {
        refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            format!("{} has no document element", surface.label()),
        )
    }
}

fn equivalent_surface_namespace(actual: &[u8], expected: &[u8]) -> bool {
    actual == expected
        || matches!(
            (actual, expected),
            (STRICT_PML, PML)
                | (STRICT_DML, DML)
                | (STRICT_CHART, CHART)
                | (STRICT_DIAGRAM, DIAGRAM)
                | (STRICT_CHART_DRAWING, CHART_DRAWING)
        )
}

fn allowed_surface_element_namespace(surface: XmlSurface, value: &[u8]) -> bool {
    match surface {
        XmlSurface::Slide => matches!(
            value,
            PML | STRICT_PML | DML | STRICT_DML | CHART | STRICT_CHART | DIAGRAM | STRICT_DIAGRAM
        ),
        XmlSurface::Chart => matches!(value, CHART | STRICT_CHART | DML | STRICT_DML),
        XmlSurface::ChartDrawing => matches!(
            value,
            CHART_DRAWING | STRICT_CHART_DRAWING | DML | STRICT_DML
        ),
        XmlSurface::DiagramData
        | XmlSurface::DiagramLayout
        | XmlSurface::DiagramStyle
        | XmlSurface::DiagramColor => {
            matches!(value, DIAGRAM | STRICT_DIAGRAM | DML | STRICT_DML)
        },
        XmlSurface::DiagramDrawing => {
            matches!(value, DIAGRAM_DRAWING | DML | STRICT_DML)
        },
        XmlSurface::Tag => matches!(value, PML | STRICT_PML),
        XmlSurface::ThemeOverride => matches!(value, DML | STRICT_DML),
        XmlSurface::ChartStyle => value == CHART_STYLE || matches!(value, DML | STRICT_DML),
        XmlSurface::ChartColor => value == CHART_COLOR || matches!(value, DML | STRICT_DML),
    }
}

fn allowed_surface_attribute_namespace(surface: XmlSurface, value: &[u8]) -> bool {
    allowed_surface_element_namespace(surface, value) || matches!(value, REL | STRICT_REL | XML)
}

fn reject_mce(xml: &[u8], context: &'static str) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MCE) {
                    return refusal(
                        SlideCopyRefusal::MarkupCompatibility,
                        format!("{context} contains markup-compatibility markup"),
                    );
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    let is_declaration = key == b"xmlns" || key.starts_with(b"xmlns:");
                    if is_declaration {
                        let value = attribute
                            .normalized_value(XmlVersion::Implicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        if value.as_bytes() == MCE {
                            return refusal(
                                SlideCopyRefusal::MarkupCompatibility,
                                format!("{context} contains a markup-compatibility declaration"),
                            );
                        }
                    } else if matches!(
                        reader.resolver().resolve_attribute(attribute.key).0,
                        ResolveResult::Bound(Namespace(value)) if value == MCE
                    ) {
                        return refusal(
                            SlideCopyRefusal::MarkupCompatibility,
                            format!("{context} contains a markup-compatibility attribute"),
                        );
                    }
                }
            },
            Event::Eof => return Ok(()),
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::UnknownSemanticSurface,
                    format!("{context} contains a document type or processing instruction"),
                );
            },
            _ => {},
        }
    }
}

fn has_signature_infrastructure(package: &OpcPackage) -> bool {
    package
        .rels()
        .iter()
        .any(|relationship| is_signature_relationship(relationship.reltype()))
        || package.iter_parts().any(|part| {
            part.partname()
                .as_str()
                .get(.."/_xmlsignatures/".len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("/_xmlsignatures/"))
                || matches!(
                    part.content_type(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN
                        | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                        | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
                )
                || part
                    .rels()
                    .iter()
                    .any(|relationship| is_signature_relationship(relationship.reltype()))
        })
}

fn is_signature_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_xml_content_type(value: &str) -> bool {
    let essence = value
        .split_once(';')
        .map_or(value, |(essence, _)| essence)
        .trim();
    essence.eq_ignore_ascii_case("application/xml")
        || essence.eq_ignore_ascii_case("text/xml")
        || essence
            .get(essence.len().saturating_sub("+xml".len())..)
            .is_some_and(|suffix| suffix.eq_ignore_ascii_case("+xml"))
}

fn available_copy_name(
    package: &OpcPackage,
    source: &PackURI,
    max_candidates: usize,
) -> Result<PackURI> {
    let value = source.as_str();
    let (stem, extension) = value
        .rfind('.')
        .map_or((value, ""), |position| value.split_at(position));
    let candidate_count = package
        .part_count()
        .checked_add(1)
        .ok_or_else(|| invalid("slide-copy candidate count overflow"))?
        .min(
            max_candidates
                .checked_add(1)
                .ok_or_else(|| invalid("slide-copy candidate limit overflow"))?,
        );
    let decimal_digits = (usize::BITS as usize)
        .checked_mul(30_103)
        .and_then(|bits| bits.checked_div(100_000))
        .and_then(|digits| digits.checked_add(1))
        .ok_or_else(|| invalid("slide-copy index digit bound overflow"))?;
    for index in 1..=candidate_count {
        let capacity = stem
            .len()
            .checked_add("-copy".len())
            .and_then(|length| length.checked_add(decimal_digits))
            .and_then(|length| length.checked_add(extension.len()))
            .ok_or_else(|| invalid("slide-copy candidate part-name length overflow"))?;
        let mut value = String::new();
        value
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "slide-copy candidate part name",
                source,
            })?;
        write!(&mut value, "{stem}-copy{index}{extension}")
            .map_err(|error| invalid(format!("cannot format slide-copy part name: {error}")))?;
        let candidate = PackURI::new(value).map_err(Error::Invalid)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    refusal(
        SlideCopyRefusal::AmbiguousTopology,
        "the deterministic copy part-name space is exhausted",
    )
}

fn validate_registered_layout(
    package: &OpcPackage,
    presentation_part: &dyn Part,
    selected: &PackURI,
) -> Result<()> {
    let presentation = crate::Presentation::new(
        crate::parts::PresentationPart::from_part(presentation_part)?,
        package,
    );
    if presentation
        .slide_layouts()?
        .iter()
        .any(|layout| layout.part().part().partname() == selected)
    {
        Ok(())
    } else {
        refusal(
            SlideCopyRefusal::SharedOwner,
            "the selected layout is not registered by a presentation slide master",
        )
    }
}

fn validate_slide_ids(slides: &[Slide]) -> Result<()> {
    if slides
        .iter()
        .any(|slide| !(MIN_SLIDE_ID..=MAX_SLIDE_ID).contains(&slide.id))
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "an existing slide ID is outside the PresentationML range",
        );
    }
    Ok(())
}

fn next_slide_id(slides: &[Slide]) -> Result<u32> {
    let maximum = slides.iter().map(|slide| slide.id).max().unwrap_or(255);
    if maximum < MAX_SLIDE_ID {
        return maximum
            .checked_add(1)
            .ok_or_else(|| invalid("slide-copy slide ID overflow"));
    }
    let mut used = Vec::new();
    used.try_reserve_exact(slides.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy used slide IDs",
            source,
        })?;
    used.extend(slides.iter().map(|slide| slide.id));
    used.sort_unstable();
    let mut candidate = MIN_SLIDE_ID;
    for value in used {
        if value == candidate {
            candidate = candidate
                .checked_add(1)
                .ok_or_else(|| invalid("slide-copy slide ID overflow"))?;
        } else if value > candidate {
            return Ok(candidate);
        }
    }
    if candidate <= MAX_SLIDE_ID {
        return Ok(candidate);
    }
    refusal(
        SlideCopyRefusal::AmbiguousTopology,
        "the PresentationML slide-ID space is exhausted",
    )
}

fn next_relationship_id(relationships: &litchi_opc::Relationships) -> Result<String> {
    let mut used = Vec::new();
    used.try_reserve_exact(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "slide-copy used presentation relationship IDs",
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
                .ok_or_else(|| Error::SlideCopyPlan {
                    kind: SlideCopyRefusal::AmbiguousTopology,
                    detail: "the presentation relationship-ID space is exhausted".to_owned(),
                })?;
        } else if value > candidate {
            break;
        }
    }
    Ok(format!("rId{candidate}"))
}

fn resolve_slide(snapshot: &Snapshot, key: crate::slide::Key<'_>) -> Result<Slide> {
    match key {
        crate::slide::Key::Index(index) => {
            snapshot
                .slides
                .get(index)
                .cloned()
                .ok_or(Error::SlideIndexOutOfBounds {
                    index,
                    len: snapshot.slides.len(),
                })
        },
        crate::slide::Key::Name(name) => {
            let mut matches = snapshot.slides.iter().filter(|slide| slide.name == name);
            let selected = matches.next().cloned();
            if matches.next().is_some() {
                return Err(Error::AmbiguousSlideName {
                    name: name.to_owned(),
                    matches: snapshot
                        .slides
                        .iter()
                        .filter(|slide| slide.name == name)
                        .count(),
                });
            }
            selected.ok_or_else(|| Error::SlideNameNotFound(name.to_owned()))
        },
    }
}

fn try_copy_string(value: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

fn refusal<T>(kind: SlideCopyRefusal, detail: impl Into<String>) -> Result<T> {
    Err(Error::SlideCopyPlan {
        kind,
        detail: detail.into(),
    })
}
