//! Bounded cross-presentation copying over deferred PresentationML sources.
//!
//! This first source-backed tranche has a deliberately bounded closure. It
//! copies one slide whose exact internal relationships are either one
//! slide-layout edge or that edge plus one direct embedded picture, and reuses
//! a destination layout only after its registered layout/master/theme boundary
//! has been proven equivalent.

use std::collections::HashSet;
use std::io::{self, Write};
use std::sync::{Arc, Mutex};

use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartView, Relationships, SourceBackedPackage, SourceLineage, SourceTopologyPlan,
    TargetMode,
};
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest, Sha256};

use super::source::{SourceBackedPresentation, SourceBackedPresentationEditor};
use crate::parts::{PresentationPart, SlideReference};
use crate::{Error, Result, SlideCopyRefusal};

const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;
const MAX_STAGING_OVERHEAD: usize = 256;
const MIN_SLIDE_ID: u32 = 256;
const MAX_SLIDE_ID: u32 = 2_147_483_647;

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_MCE_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/markup-compatibility/2006";
const TRANSITIONAL_REL_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_PML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const TRANSITIONAL_DML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const P14_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const STRICT_SLIDE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const STRICT_LAYOUT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";
const STRICT_IMAGE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const STRICT_MASTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_THEME_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

#[derive(Clone, Copy)]
enum RootNamespace {
    PresentationMl,
    DrawingMl,
}

impl RootNamespace {
    const fn matches(self, pml: bool, dml: bool) -> bool {
        match self {
            Self::PresentationMl => pml,
            Self::DrawingMl => dml,
        }
    }
}

/// An opaque, one-way source-backed cross-presentation slide-copy plan.
///
/// Only a dependency-free slide closure is supported in this tranche: the
/// source slide has exactly one internal slide-layout relationship, optionally
/// plus one direct embedded image, and no chart, diagram, table, notes,
/// comments, external, or shared-owner relationship. No inverse or durable
/// patch is represented.
pub struct SourceBackedCrossSlideCopyPlan {
    source: SourceBackedPresentation,
    source_position: usize,
    destination_slide_position: usize,
    insertion_position: usize,
    destination_slide_count: usize,
    source_name: String,
    source_version: SourceVersion,
    destination_version: SourceVersion,
    source_lineage: SourceLineage,
    destination_lineage: SourceLineage,
    source_slide_uri: PackURI,
    destination_layout_uri: PackURI,
    presentation_uri: PackURI,
    target_slide_uri: PackURI,
    slide_id: u32,
    presentation_relationship_id: String,
    layout_relationship_id: String,
    slide_relationship_type: String,
    layout_relationship_type: String,
    source_slide_xml: Vec<u8>,
    target_presentation_xml: Vec<u8>,
    image: Option<PreparedImage>,
    touched_digest: [u8; 32],
    planned_bytes: usize,
    _memory_reservation: Option<Arc<Reservation>>,
}

#[derive(PartialEq, Eq)]
struct PreparedImage {
    source_uri: PackURI,
    target_uri: PackURI,
    relationship_id: String,
    relationship_type: String,
    content_type: String,
    bytes: Vec<u8>,
}

/// Semantic result of one published source-backed cross-presentation copy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceBackedCrossSlideCopySnapshot {
    source_position: usize,
    destination_slide_position: usize,
    insertion_position: usize,
    destination_slide_count: usize,
    source_name: String,
}

impl SourceBackedCrossSlideCopyPlan {
    /// Source zero-based slide position captured by this plan.
    #[must_use]
    pub const fn source_position(&self) -> usize {
        self.source_position
    }

    /// Destination slide position whose registered layout is reused.
    #[must_use]
    pub const fn destination_slide_position(&self) -> usize {
        self.destination_slide_position
    }

    /// Destination zero-based insertion position captured by this plan.
    #[must_use]
    pub const fn insertion_position(&self) -> usize {
        self.insertion_position
    }

    /// Destination slide count after the planned insertion.
    #[must_use]
    pub const fn destination_slide_count(&self) -> usize {
        self.destination_slide_count
    }
}

impl SourceBackedCrossSlideCopySnapshot {
    /// Source zero-based slide position published by this operation.
    #[must_use]
    pub const fn source_position(&self) -> usize {
        self.source_position
    }

    /// Destination slide position whose registered layout was reused.
    #[must_use]
    pub const fn destination_slide_position(&self) -> usize {
        self.destination_slide_position
    }

    /// Destination zero-based insertion position used by this operation.
    #[must_use]
    pub const fn insertion_position(&self) -> usize {
        self.insertion_position
    }

    /// Destination slide count after publication.
    #[must_use]
    pub const fn destination_slide_count(&self) -> usize {
        self.destination_slide_count
    }

    /// Producer-visible name retained by the copied slide.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.source_name
    }
}

impl SourceBackedPresentationEditor {
    /// Plan one dependency-free source-backed slide copy.
    ///
    /// `source_position` selects the source slide, `destination_slide_position`
    /// selects an existing destination slide whose registered layout is reused,
    /// and `insertion_position` selects the destination slide-list position.
    pub fn plan_cross_slide_copy(
        &self,
        source: &SourceBackedPresentation,
        source_position: usize,
        destination_slide_position: usize,
        insertion_position: usize,
    ) -> Result<SourceBackedCrossSlideCopyPlan> {
        let prepared = prepare(
            self,
            source,
            source_position,
            destination_slide_position,
            insertion_position,
        )?;
        Ok(SourceBackedCrossSlideCopyPlan {
            source: source.clone(),
            source_position,
            destination_slide_position,
            insertion_position,
            destination_slide_count: prepared.destination_slide_count,
            source_name: prepared.source_name,
            source_version: prepared.source_version,
            destination_version: prepared.destination_version,
            source_lineage: prepared.source_lineage,
            destination_lineage: prepared.destination_lineage,
            source_slide_uri: prepared.source_slide_uri,
            destination_layout_uri: prepared.destination_layout_uri,
            presentation_uri: prepared.presentation_uri,
            target_slide_uri: prepared.target_slide_uri,
            slide_id: prepared.slide_id,
            presentation_relationship_id: prepared.presentation_relationship_id,
            layout_relationship_id: prepared.layout_relationship_id,
            slide_relationship_type: prepared.slide_relationship_type,
            layout_relationship_type: prepared.layout_relationship_type,
            source_slide_xml: prepared.source_slide_xml,
            target_presentation_xml: prepared.target_presentation_xml,
            image: prepared.image,
            touched_digest: prepared.touched_digest,
            planned_bytes: prepared.planned_bytes,
            _memory_reservation: prepared.memory_reservation,
        })
    }

    /// Publish a previously planned one-way source-backed slide copy.
    ///
    /// The bounded planner is rerun against both current lineages and
    /// revisions before output. The OPC topology publisher owns raw ZIP
    /// preservation, content-types, canonical relationship members, ZIP32
    /// limits, and incomplete-output classification.
    pub fn publish_cross_slide_copy_to_stream<W: Write>(
        self,
        writer: W,
        plan: &SourceBackedCrossSlideCopyPlan,
    ) -> Result<SourceBackedCrossSlideCopySnapshot> {
        self.package.check_execution()?;
        plan.source.check_source()?;
        if self.package.source_lineage() != plan.destination_lineage
            || plan.source.inner.package.source_lineage() != plan.source_lineage
        {
            return Err(Error::StaleSource);
        }
        let current = prepare(
            &self,
            &plan.source,
            plan.source_position,
            plan.destination_slide_position,
            plan.insertion_position,
        )?;
        if !current.matches(plan) {
            return Err(Error::StaleSource);
        }
        verify_candidate(&self, &plan.source, &current)?;
        let _memory_reservation = current.memory_reservation.clone();
        let topology = current.into_topology()?;
        let source_check_state = Arc::new(Mutex::new(SourceCheckState::default()));
        let writer = SourceCheckedWriter {
            inner: writer,
            source: plan.source.clone(),
            state: Arc::clone(&source_check_state),
        };
        let publication = self.package.write_topology_to_stream(writer, topology);
        if let Some(source) = take_source_failure(&source_check_state) {
            let written = accepted_bytes(&source_check_state);
            return Err(source_failure_with_progress(written, source));
        }
        publication?;
        if let Err(error) = plan.source.check_source() {
            let written = accepted_bytes(&source_check_state);
            return Err(match error {
                Error::Opc(source) => source_failure_with_progress(written, source),
                other => other,
            });
        }
        Ok(SourceBackedCrossSlideCopySnapshot {
            source_position: plan.source_position,
            destination_slide_position: plan.destination_slide_position,
            insertion_position: plan.insertion_position,
            destination_slide_count: plan.destination_slide_count,
            source_name: plan.source_name.clone(),
        })
    }
}

struct Prepared {
    destination_slide_position: usize,
    destination_slide_count: usize,
    insertion_position: usize,
    source_name: String,
    source_version: SourceVersion,
    destination_version: SourceVersion,
    source_lineage: SourceLineage,
    destination_lineage: SourceLineage,
    source_slide_uri: PackURI,
    destination_layout_uri: PackURI,
    presentation_uri: PackURI,
    target_slide_uri: PackURI,
    slide_id: u32,
    presentation_relationship_id: String,
    layout_relationship_id: String,
    slide_relationship_type: String,
    layout_relationship_type: String,
    source_slide_xml: Vec<u8>,
    target_presentation_xml: Vec<u8>,
    image: Option<PreparedImage>,
    touched_digest: [u8; 32],
    planned_bytes: usize,
    memory_reservation: Option<Arc<Reservation>>,
}

impl Prepared {
    fn matches(&self, plan: &SourceBackedCrossSlideCopyPlan) -> bool {
        self.destination_slide_count == plan.destination_slide_count
            && self.destination_slide_position == plan.destination_slide_position
            && self.insertion_position == plan.insertion_position
            && self.source_name == plan.source_name
            && self.source_version == plan.source_version
            && self.destination_version == plan.destination_version
            && self.source_lineage == plan.source_lineage
            && self.destination_lineage == plan.destination_lineage
            && self.source_slide_uri == plan.source_slide_uri
            && self.destination_layout_uri == plan.destination_layout_uri
            && self.presentation_uri == plan.presentation_uri
            && self.target_slide_uri == plan.target_slide_uri
            && self.slide_id == plan.slide_id
            && self.presentation_relationship_id == plan.presentation_relationship_id
            && self.layout_relationship_id == plan.layout_relationship_id
            && self.slide_relationship_type == plan.slide_relationship_type
            && self.layout_relationship_type == plan.layout_relationship_type
            && self.source_slide_xml == plan.source_slide_xml
            && self.target_presentation_xml == plan.target_presentation_xml
            && self.image == plan.image
            && self.touched_digest == plan.touched_digest
            && self.planned_bytes == plan.planned_bytes
    }

    fn into_topology(self) -> Result<SourceTopologyPlan> {
        let mut topology = SourceTopologyPlan::new();
        topology.try_replace_part(self.presentation_uri.clone(), self.target_presentation_xml)?;
        topology.try_add_part(
            self.target_slide_uri.clone(),
            ct::PML_SLIDE,
            self.source_slide_xml,
        )?;
        topology.try_add_internal_relationship(
            self.presentation_uri,
            self.presentation_relationship_id,
            self.slide_relationship_type,
            self.target_slide_uri.clone(),
        )?;
        topology.try_add_internal_relationship(
            self.target_slide_uri.clone(),
            self.layout_relationship_id,
            self.layout_relationship_type,
            self.destination_layout_uri,
        )?;
        if let Some(image) = self.image {
            topology.try_add_part(image.target_uri.clone(), image.content_type, image.bytes)?;
            topology.try_add_internal_relationship(
                self.target_slide_uri.clone(),
                image.relationship_id,
                image.relationship_type,
                image.target_uri,
            )?;
        }
        Ok(topology)
    }
}

fn prepare(
    editor: &SourceBackedPresentationEditor,
    source: &SourceBackedPresentation,
    source_position: usize,
    destination_slide_position: usize,
    insertion_position: usize,
) -> Result<Prepared> {
    editor.package.check_execution()?;
    source.check_source()?;
    let source_lineage = source.inner.package.source_lineage();
    let destination_lineage = editor.package.source_lineage();
    let source_version = source.inner.package.source_version()?;
    let destination_version = editor.package.source_version()?;
    if source_lineage == destination_lineage || source_version == destination_version {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "source and destination share one source-backed lineage or revision",
        );
    }

    let source_presentation_view = source.inner.package.main_document_part()?;
    let source_presentation_data = source_presentation_view.data()?;
    let destination_presentation_view = editor.package.main_document_part()?;
    let destination_presentation_data = destination_presentation_view.data()?;
    source.inner.package.validate_topology_source_boundary()?;
    editor.package.validate_topology_source_boundary()?;
    reject_package_features(
        &source.inner.package,
        source_presentation_data.as_bytes(),
        "source",
    )?;
    reject_package_features(
        &editor.package,
        destination_presentation_data.as_bytes(),
        "destination",
    )?;

    let source_presentation = super::source::SourcePart::from_view(
        &source_presentation_view,
        source_presentation_data.clone(),
    );
    let destination_presentation = super::source::SourcePart::from_view(
        &destination_presentation_view,
        destination_presentation_data.clone(),
    );
    let source_main = PresentationPart::from_part(&source_presentation)?;
    let destination_main = PresentationPart::from_part(&destination_presentation)?;
    let source_dialect = validate_xml(
        source_presentation_data.as_bytes(),
        b"presentation",
        RootNamespace::PresentationMl,
        true,
        true,
        true,
        true,
        false,
        "source presentation",
    )?;
    let destination_dialect = validate_xml(
        destination_presentation_data.as_bytes(),
        b"presentation",
        RootNamespace::PresentationMl,
        true,
        true,
        true,
        true,
        false,
        "destination presentation",
    )?;
    if source_dialect != destination_dialect {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "source and destination use different OOXML dialects",
        );
    }
    validate_presentation_relationships(source_presentation_view.rels(), "source")?;
    validate_presentation_relationships(destination_presentation_view.rels(), "destination")?;

    let source_refs = source_main.slide_references()?;
    let destination_refs = destination_main.slide_references()?;
    validate_slide_refs(&source_refs, "source")?;
    validate_slide_refs(&destination_refs, "destination")?;
    if source_refs.len() != source.inner.slides.len()
        || destination_refs.len() != editor.slides.len()
    {
        return Err(Error::StaleSource);
    }
    for (reference, slide) in source_refs.iter().zip(source.inner.slides.iter()) {
        if reference.id() != slide.slide_id
            || reference.relationship_id() != slide.binding.slide_reference_id
        {
            return Err(Error::StaleSource);
        }
    }
    for (reference, slide) in destination_refs.iter().zip(editor.slides.iter()) {
        if reference.id() != slide.slide_id
            || reference.relationship_id() != slide.binding.slide_reference_id
        {
            return Err(Error::StaleSource);
        }
    }
    validate_registered_slide_parts(&source.inner.package, &source.inner.slides, "source")?;
    validate_registered_slide_parts(&editor.package, &editor.slides, "destination")?;
    if source_position >= source_refs.len() {
        return Err(Error::SlideIndexOutOfBounds {
            index: source_position,
            len: source_refs.len(),
        });
    }
    if destination_slide_position >= destination_refs.len() {
        return Err(Error::SlideIndexOutOfBounds {
            index: destination_slide_position,
            len: destination_refs.len(),
        });
    }
    if insertion_position > destination_refs.len() {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "destination insertion position is outside the slide list",
        );
    }

    let source_slide = source.inner.slides[source_position].clone();
    let source_slide_view = source.inner.package.part(&source_slide.part_uri)?;
    let source_slide_data = source_slide_view.data()?;
    let slide_dialect = validate_xml(
        source_slide_data.as_bytes(),
        b"sld",
        RootNamespace::PresentationMl,
        true,
        true,
        false,
        true,
        false,
        "source slide",
    )?;
    if slide_dialect != source_dialect {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "source slide and presentation use different OOXML dialects",
        );
    }
    let embedded_image = direct_embedded_image(source_slide_data.as_bytes(), slide_dialect)?;
    let layout_relationship = exact_layout_relationship(
        &source_slide_view,
        slide_dialect,
        embedded_image
            .as_ref()
            .map(|image| image.relationship_id.as_str()),
    )?;
    let source_layout_uri = layout_relationship.target_partname()?;
    let source_image = if let Some(image) = embedded_image.as_ref() {
        let relationship = source_slide_view
            .rels()
            .get(&image.relationship_id)
            .ok_or_else(|| Error::Relationship("embedded image relationship is missing".into()))?;
        let source_image_uri = relationship.target_partname()?;
        if !source_image_uri.as_str().starts_with("/ppt/media/") {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                "embedded image target is outside /ppt/media/",
            );
        }
        let source_image_part = source
            .inner
            .package
            .part(&source_image_uri)
            .map_err(|error| match error {
                litchi_opc::OpcError::PartNotFound(_) => Error::SlideCopyPlan {
                    kind: SlideCopyRefusal::AmbiguousTopology,
                    detail: "embedded image target part is missing".into(),
                },
                other => Error::Opc(other),
            })?;
        if !source_image_part.content_type().starts_with("image/") {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "embedded image target does not have an image content type",
            );
        }
        if !source_image_part.rels().is_empty() {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                "embedded image target must be a leaf part",
            );
        }
        let declared_size = source_image_part.declared_uncompressed_size()?;
        Some((
            source_image_uri,
            relationship.r_id().to_owned(),
            relationship.reltype().to_owned(),
            source_image_part.content_type().to_owned(),
            source_image_part,
            declared_size,
        ))
    } else {
        None
    };
    let source_layouts = registered_layouts(
        &source.inner.package,
        &source_presentation_view,
        source_presentation_data.as_bytes(),
        source_dialect,
        "source",
    )?;
    if !source_layouts
        .iter()
        .any(|uri| uri.is_equivalent_to(&source_layout_uri))
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "source slide layout is not registered by a slide master",
        );
    }
    let destination_layouts = registered_layouts(
        &editor.package,
        &destination_presentation_view,
        destination_presentation_data.as_bytes(),
        destination_dialect,
        "destination",
    )?;
    let destination_anchor = editor.slides[destination_slide_position].clone();
    let destination_anchor_view = editor.package.part(&destination_anchor.part_uri)?;
    let destination_anchor_data = destination_anchor_view.data()?;
    let destination_anchor_dialect = validate_xml(
        destination_anchor_data.as_bytes(),
        b"sld",
        RootNamespace::PresentationMl,
        true,
        true,
        false,
        true,
        false,
        "destination anchor slide",
    )?;
    if destination_anchor_dialect != destination_dialect {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "destination anchor slide and presentation use different OOXML dialects",
        );
    }
    let destination_layout_relationship =
        exact_layout_relationship(&destination_anchor_view, destination_anchor_dialect, None)?;
    let destination_layout_uri = destination_layout_relationship.target_partname()?;
    if !destination_layouts
        .iter()
        .any(|uri| uri.is_equivalent_to(&destination_layout_uri))
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "destination anchor layout is not registered by a slide master",
        );
    }
    let destination_anchor_reference = &destination_refs[destination_slide_position];
    let destination_anchor_relationship = destination_presentation_view
        .rels()
        .get(destination_anchor_reference.relationship_id())
        .ok_or_else(|| Error::Relationship("destination anchor relationship is missing".into()))?;
    validate_internal(
        destination_anchor_relationship,
        destination_dialect,
        rt::SLIDE,
        STRICT_SLIDE_REL,
        "destination anchor slide",
    )?;
    if !destination_anchor_relationship
        .target_partname()?
        .is_equivalent_to(&destination_anchor.part_uri)
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "destination anchor relationship targets a different slide",
        );
    }
    let graph_digest = prove_graphs(
        &source.inner.package,
        &source_layout_uri,
        &editor.package,
        &destination_layout_uri,
        source_dialect,
    )?;

    let names = destination_names(
        &editor.package,
        &destination_presentation_view,
        &destination_refs,
    )?;
    let source_name = crate::namespace::presentation_name(source_slide_data.as_bytes())?
        .filter(|name| !name.is_empty())
        .ok_or_else(|| Error::SlideCopyPlan {
            kind: SlideCopyRefusal::AmbiguousTopology,
            detail: "source slide name is missing or empty".into(),
        })?;
    if names.contains(&normalize_name(&source_name)?) {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "source slide name collides with a destination slide name",
        );
    }
    let slide_id = allocate_slide_id(&destination_refs)?;
    let presentation_relationship_id =
        allocate_relationship_id(destination_presentation_view.rels())?;
    let target_slide_uri = allocate_slide_uri(&editor.package)?;
    let target_image_uri = source_image
        .as_ref()
        .map(|(source_uri, ..)| allocate_media_uri(&editor.package, source_uri))
        .transpose()?;
    let presentation_relationship_type = destination_anchor_relationship.reltype().to_owned();
    let layout_relationship_type = destination_layout_relationship.reltype().to_owned();
    if source_image
        .as_ref()
        .is_some_and(|(_, relationship_id, ..)| {
            relationship_id == destination_layout_relationship.r_id()
        })
    {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "embedded image relationship ID collides with the destination layout relationship",
        );
    }
    let image_size =
        source_image
            .as_ref()
            .map_or(Ok(0usize), |(_, _, _, _, _, declared_size)| {
                if *declared_size > editor.limits.max_total_part_bytes() {
                    return Err(Error::Limit {
                        resource: "source-backed embedded image bytes",
                        limit: usize::try_from(editor.limits.max_total_part_bytes())
                            .unwrap_or(usize::MAX),
                    });
                }
                usize::try_from(*declared_size).map_err(|_| Error::Limit {
                    resource: "source-backed embedded image bytes",
                    limit: usize::MAX,
                })
            })?;
    let staging_request = source_slide_data
        .as_bytes()
        .len()
        .checked_add(destination_presentation_data.as_bytes().len())
        .and_then(|bytes| bytes.checked_add(image_size))
        .ok_or_else(|| invalid("source-backed cross-copy staging size overflow"))?;
    let memory_reservation = reserve_memory(
        editor.package.execution_context().as_ref(),
        staging_request,
        editor.limits,
    )?;
    let source_image_data =
        if let Some((_, _, _, _, source_image_part, declared_size)) = source_image.as_ref() {
            let data = source_image_part.data()?;
            let actual_size = data.as_bytes().len();
            let declared_size = usize::try_from(*declared_size).map_err(|_| Error::Limit {
                resource: "source-backed embedded image bytes",
                limit: usize::MAX,
            })?;
            if actual_size != declared_size {
                return Err(Error::Invalid(
                    "embedded image decoded size differs from ZIP metadata".into(),
                ));
            }
            Some(data)
        } else {
            None
        };
    let target_presentation_xml = crate::opened::insert_slide_binding(
        destination_presentation_data.as_bytes(),
        destination_refs
            .iter()
            .map(|reference| (reference.id(), reference.relationship_id())),
        insertion_position,
        slide_id,
        &presentation_relationship_id,
    )?;
    if validate_xml(
        &target_presentation_xml,
        b"presentation",
        RootNamespace::PresentationMl,
        true,
        true,
        true,
        true,
        false,
        "staged destination presentation",
    )? != destination_dialect
    {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "staged destination presentation changed OOXML dialect",
        );
    }
    validate_candidate_bindings(
        &target_presentation_xml,
        destination_refs
            .iter()
            .map(|reference| (reference.id(), reference.relationship_id())),
        insertion_position,
        slide_id,
        &presentation_relationship_id,
    )?;
    let source_slide_xml = clone_bytes(source_slide_data.as_bytes(), "source-backed copied slide")?;
    let image = if let Some((source_uri, relationship_id, relationship_type, content_type, _, _)) =
        source_image
    {
        let data = source_image_data
            .as_ref()
            .ok_or_else(|| invalid("source-backed embedded image payload is missing"))?;
        Some(PreparedImage {
            source_uri,
            target_uri: target_image_uri
                .ok_or_else(|| invalid("source-backed embedded image target URI is missing"))?,
            relationship_id,
            relationship_type,
            content_type,
            bytes: clone_bytes(data.as_bytes(), "source-backed copied image")?,
        })
    } else {
        None
    };
    let planned_bytes = source_slide_xml
        .len()
        .checked_add(target_presentation_xml.len())
        .and_then(|bytes| bytes.checked_add(image.as_ref().map_or(0, |image| image.bytes.len())))
        .ok_or_else(|| invalid("source-backed cross-copy staged byte count overflow"))?;
    let touched_digest = digest_touched(
        source_presentation_data.as_bytes(),
        destination_presentation_data.as_bytes(),
        source_slide_data.as_bytes(),
        &target_presentation_xml,
        graph_digest,
        &source_slide.part_uri,
        &source_layout_uri,
        &destination_layout_uri,
        &target_slide_uri,
        slide_id,
        &presentation_relationship_id,
        &source_name,
        &names,
        image.as_ref(),
    )?;
    Ok(Prepared {
        destination_slide_position,
        destination_slide_count: destination_refs.len() + 1,
        insertion_position,
        source_name,
        source_version,
        destination_version,
        source_lineage,
        destination_lineage,
        source_slide_uri: source_slide.part_uri.clone(),
        destination_layout_uri,
        presentation_uri: destination_presentation_view.partname().clone(),
        target_slide_uri,
        slide_id,
        presentation_relationship_id,
        layout_relationship_id: destination_layout_relationship.r_id().to_owned(),
        slide_relationship_type: presentation_relationship_type,
        layout_relationship_type: layout_relationship_type.to_owned(),
        source_slide_xml,
        target_presentation_xml,
        image,
        touched_digest,
        planned_bytes,
        memory_reservation,
    })
}

fn verify_candidate(
    editor: &SourceBackedPresentationEditor,
    source: &SourceBackedPresentation,
    prepared: &Prepared,
) -> Result<()> {
    let destination = editor.package.part(&prepared.presentation_uri)?;
    if destination.content_type() == ct::PML_PRES_MACRO_MAIN
        || destination.content_type() == ct::PML_SLIDESHOW_MACRO_MAIN
        || destination.content_type() == ct::PML_TEMPLATE_MACRO_MAIN
    {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            "macro-enabled destination presentation is unsupported",
        );
    }
    validate_candidate_bindings(
        &prepared.target_presentation_xml,
        editor
            .slides
            .iter()
            .map(|slide| (slide.slide_id, slide.binding.slide_reference_id.as_str())),
        prepared.insertion_position,
        prepared.slide_id,
        &prepared.presentation_relationship_id,
    )?;
    let source_view = source.inner.package.part(&prepared.source_slide_uri)?;
    let source_data = source_view.data()?;
    if source_data.as_bytes() != prepared.source_slide_xml.as_slice() {
        return Err(Error::StaleSource);
    }
    if let Some(image) = &prepared.image {
        let image_view = source.inner.package.part(&image.source_uri)?;
        let image_data = image_view.data()?;
        if image_data.as_bytes() != image.bytes.as_slice()
            || image_view.content_type() != image.content_type.as_str()
            || !image_view.rels().is_empty()
        {
            return Err(Error::StaleSource);
        }
    }
    let staged_source_dialect = validate_xml(
        prepared.source_slide_xml.as_slice(),
        b"sld",
        RootNamespace::PresentationMl,
        true,
        true,
        false,
        true,
        false,
        "staged source slide",
    )?;
    let staged_destination_dialect = validate_xml(
        &prepared.target_presentation_xml,
        b"presentation",
        RootNamespace::PresentationMl,
        true,
        true,
        true,
        true,
        false,
        "staged destination presentation",
    )?;
    if staged_source_dialect != staged_destination_dialect {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "staged source and destination use different OOXML dialects",
        );
    }
    if part_exists(&editor.package, &prepared.target_slide_uri)?
        || editor.package.non_part_members().iter().any(|member| {
            member
                .name()
                .eq_ignore_ascii_case(prepared.target_slide_uri.membername())
        })
    {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            "allocated destination slide member already exists",
        );
    }
    if let Some(image) = &prepared.image {
        let relationship_uri = image
            .target_uri
            .rels_uri()
            .map_err(|error| Error::Uri(error.to_string()))?;
        if part_exists(&editor.package, &image.target_uri)?
            || editor.package.non_part_members().iter().any(|member| {
                member
                    .name()
                    .eq_ignore_ascii_case(image.target_uri.membername())
                    || member
                        .name()
                        .eq_ignore_ascii_case(relationship_uri.membername())
            })
        {
            return refusal(
                SlideCopyRefusal::UnknownPhysicalMember,
                "allocated destination image member already exists",
            );
        }
    }
    Ok(())
}

fn part_exists(package: &SourceBackedPackage, uri: &PackURI) -> Result<bool> {
    match package.part(uri) {
        Ok(_) => Ok(true),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(false),
        Err(error) => Err(Error::Opc(error)),
    }
}

fn exact_layout_relationship<'a>(
    slide: &'a PartView<'a>,
    dialect: Dialect,
    image_relationship_id: Option<&str>,
) -> Result<&'a litchi_opc::Relationship> {
    let expected_relationships = 1usize
        .checked_add(usize::from(image_relationship_id.is_some()))
        .ok_or_else(|| invalid("source slide relationship count overflow"))?;
    if slide.rels().len() != expected_relationships {
        return refusal(
            SlideCopyRefusal::UnsupportedRelationship,
            "source slide has unsupported or extra relationships",
        );
    }
    let mut layout_relationship = None;
    let mut image_found = false;
    for relationship in slide.rels().iter() {
        if relation_matches(
            relationship.reltype(),
            dialect,
            rt::SLIDE_LAYOUT,
            STRICT_LAYOUT_REL,
        ) {
            if layout_relationship.is_some() {
                return refusal(
                    SlideCopyRefusal::AmbiguousTopology,
                    "source slide has multiple slide-layout relationships",
                );
            }
            validate_internal(
                relationship,
                dialect,
                rt::SLIDE_LAYOUT,
                STRICT_LAYOUT_REL,
                "slide layout",
            )?;
            layout_relationship = Some(relationship);
        } else if image_relationship_id.is_some_and(|id| id == relationship.r_id()) {
            validate_internal(
                relationship,
                dialect,
                rt::IMAGE,
                STRICT_IMAGE_REL,
                "embedded image",
            )?;
            image_found = true;
        } else {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                "source slide has an unsupported relationship",
            );
        }
    }
    if image_relationship_id.is_some() && !image_found {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "embedded image relationship is missing",
        );
    }
    layout_relationship.ok_or_else(|| Error::Relationship("source slide layout missing".into()))
}

struct EmbeddedImageReference {
    relationship_id: String,
}

fn direct_embedded_image(xml: &[u8], dialect: Dialect) -> Result<Option<EmbeddedImageReference>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut scene_depth = None;
    let mut picture_depth = None;
    let mut picture_fill_depth = None;
    let mut scene_trees = 0usize;
    let mut picture_blip_fill_seen = false;
    let mut pictures = 0usize;
    let mut blips = 0usize;
    let mut relationship_id = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml = is_pml(&namespace);
        let dml = is_dml(&namespace);
        let is_start = matches!(&event, Event::Start(_));
        let is_empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local = element.local_name();
                if pml && local.as_ref() == b"spTree" {
                    scene_trees = scene_trees
                        .checked_add(1)
                        .ok_or_else(|| invalid("source slide shape-tree count overflow"))?;
                    if scene_trees > 1 || scene_depth.is_some() {
                        return refusal(
                            SlideCopyRefusal::AmbiguousTopology,
                            "source slide must contain exactly one shape tree",
                        );
                    }
                    if is_empty {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "source slide shape tree cannot be empty",
                        );
                    }
                    scene_depth = Some(
                        depth
                            .checked_add(1)
                            .ok_or_else(|| invalid("source slide XML depth overflow"))?,
                    );
                } else if pml && local.as_ref() == b"pic" {
                    if scene_depth != Some(depth) {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "source slide picture is not a direct scene element",
                        );
                    }
                    pictures = pictures
                        .checked_add(1)
                        .ok_or_else(|| invalid("source slide picture count overflow"))?;
                    if pictures > 1 {
                        return refusal(
                            SlideCopyRefusal::AmbiguousTopology,
                            "source slide contains multiple direct pictures",
                        );
                    }
                    if is_empty {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "source slide picture has no picture content",
                        );
                    }
                    picture_depth = Some(
                        depth
                            .checked_add(1)
                            .ok_or_else(|| invalid("source slide XML depth overflow"))?,
                    );
                    picture_fill_depth = None;
                    picture_blip_fill_seen = false;
                } else if pml && local.as_ref() == b"blipFill" {
                    if picture_depth.is_some() {
                        if picture_depth != Some(depth) || picture_blip_fill_seen {
                            return refusal(
                                SlideCopyRefusal::UnknownSemanticSurface,
                                "source slide picture has an unsupported blipFill structure",
                            );
                        }
                        if is_empty {
                            return refusal(
                                SlideCopyRefusal::UnknownSemanticSurface,
                                "source slide picture blipFill cannot be empty",
                            );
                        }
                        picture_blip_fill_seen = true;
                        picture_fill_depth = Some(
                            depth
                                .checked_add(1)
                                .ok_or_else(|| invalid("source slide XML depth overflow"))?,
                        );
                    }
                } else if dml && local.as_ref() == b"blip" {
                    if picture_depth.is_none() || picture_fill_depth != Some(depth) {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "source slide image is outside p:pic/p:blipFill",
                        );
                    }
                    blips = blips
                        .checked_add(1)
                        .ok_or_else(|| invalid("source slide image count overflow"))?;
                    if blips > 1 {
                        return refusal(
                            SlideCopyRefusal::AmbiguousTopology,
                            "source slide picture contains multiple images",
                        );
                    }
                    let expected_namespace = match dialect {
                        Dialect::Transitional => TRANSITIONAL_REL_NAMESPACE,
                        Dialect::Strict => STRICT_REL_NAMESPACE,
                    };
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let key = attribute.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            continue;
                        }
                        let attribute_namespace =
                            reader.resolver().resolve_attribute(attribute.key).0;
                        let relationship_namespace = match attribute_namespace {
                            ResolveResult::Bound(Namespace(value))
                                if value == TRANSITIONAL_REL_NAMESPACE
                                    || value == STRICT_REL_NAMESPACE =>
                            {
                                Some(value)
                            },
                            _ => None,
                        };
                        if key.starts_with(b"r:") || relationship_namespace.is_some() {
                            if relationship_namespace != Some(expected_namespace) {
                                return refusal(
                                    SlideCopyRefusal::UnsupportedRelationship,
                                    "source slide picture uses a mixed relationship namespace",
                                );
                            }
                            match attribute.key.local_name().as_ref() {
                                b"embed" => {
                                    if relationship_id.is_some() || attribute.value.is_empty() {
                                        return refusal(
                                            SlideCopyRefusal::AmbiguousTopology,
                                            "source slide picture has multiple or empty image IDs",
                                        );
                                    }
                                    relationship_id = Some(
                                        std::str::from_utf8(attribute.value.as_ref())
                                            .map_err(|_| {
                                                Error::Invalid(
                                                    "source slide image relationship ID is not UTF-8"
                                                        .into(),
                                                )
                                            })?
                                            .to_owned(),
                                    );
                                },
                                b"link" => {
                                    return refusal(
                                        SlideCopyRefusal::UnsupportedRelationship,
                                        "linked images are unsupported",
                                    );
                                },
                                _ => {
                                    return refusal(
                                        SlideCopyRefusal::UnsupportedRelationship,
                                        "source slide picture has an unsupported relationship attribute",
                                    );
                                },
                            }
                        }
                    }
                }
                if is_start {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("source slide XML depth overflow"))?;
                }
            },
            Event::End(_) => {
                let closing_depth = depth;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("source slide XML depth underflow"))?;
                if picture_depth == Some(closing_depth) {
                    if !picture_blip_fill_seen {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            "source slide picture has no direct blipFill",
                        );
                    }
                    picture_depth = None;
                }
                if picture_fill_depth == Some(closing_depth) {
                    picture_fill_depth = None;
                }
                if scene_depth == Some(closing_depth) {
                    scene_depth = None;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::MarkupCompatibility,
                    "source slide contains DTD or PI",
                );
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0
        || scene_depth.is_some()
        || picture_depth.is_some()
        || picture_fill_depth.is_some()
    {
        return Err(invalid("source slide XML is unterminated"));
    }
    if scene_trees != 1 {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "source slide must contain exactly one shape tree",
        );
    }
    if pictures == 0 {
        return Ok(None);
    }
    let relationship_id = relationship_id.ok_or_else(|| Error::SlideCopyPlan {
        kind: SlideCopyRefusal::UnknownSemanticSurface,
        detail: "source slide picture has no embedded image relationship".into(),
    })?;
    if blips != 1 {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            "source slide picture has invalid image content",
        );
    }
    Ok(Some(EmbeddedImageReference { relationship_id }))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Dialect {
    Transitional,
    Strict,
}

fn registered_layouts(
    package: &SourceBackedPackage,
    presentation: &PartView<'_>,
    xml: &[u8],
    dialect: Dialect,
    context: &'static str,
) -> Result<Vec<PackURI>> {
    let master_ids = relationship_list(xml, b"sldMasterIdLst", b"sldMasterId")?;
    if master_ids.is_empty() {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            format!("{context} has no slide masters"),
        );
    }
    let mut result = Vec::new();
    let mut master_seen = HashSet::new();
    let mut layout_seen = HashSet::new();
    result
        .try_reserve_exact(package.iter_parts().count())
        .map_err(|source| Error::Allocation {
            resource: "source-backed registered slide layouts",
            source,
        })?;
    master_seen
        .try_reserve(master_ids.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed registered slide masters",
            source,
        })?;
    layout_seen
        .try_reserve(package.iter_parts().count())
        .map_err(|source| Error::Allocation {
            resource: "source-backed registered slide-layout index",
            source,
        })?;
    for id in master_ids {
        let relationship = presentation.rels().get(&id).ok_or_else(|| {
            Error::Relationship(format!(
                "{context} slide master relationship '{id}' missing"
            ))
        })?;
        validate_internal(
            relationship,
            dialect,
            rt::SLIDE_MASTER,
            STRICT_MASTER_REL,
            "slide master",
        )?;
        let master_uri = relationship.target_partname()?;
        if !master_seen.insert(folded_ascii_name(
            master_uri.as_str(),
            "source-backed registered slide-master name",
        )?) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{context} repeats a slide master"),
            );
        }
        let master = package.part(&master_uri)?;
        if master.content_type() != ct::PML_SLIDE_MASTER {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE_MASTER.into(),
                actual: master.content_type().into(),
            });
        }
        let master_data = master.data()?;
        if validate_xml(
            master_data.as_bytes(),
            b"sldMaster",
            RootNamespace::PresentationMl,
            true,
            true,
            true,
            false,
            false,
            "slide master",
        )? != dialect
        {
            return refusal(
                SlideCopyRefusal::UnknownSemanticSurface,
                format!("{context} slide master mixes dialects"),
            );
        }
        let layout_ids =
            relationship_list(master_data.as_bytes(), b"sldLayoutIdLst", b"sldLayoutId")?;
        let mut relation_count = 0usize;
        for relationship in master.rels().iter() {
            if relation_matches(
                relationship.reltype(),
                dialect,
                rt::SLIDE_LAYOUT,
                STRICT_LAYOUT_REL,
            ) || relation_matches(relationship.reltype(), dialect, rt::THEME, STRICT_THEME_REL)
            {
                relation_count = relation_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("relationship count overflow"))?;
            } else {
                return refusal(
                    SlideCopyRefusal::UnsupportedRelationship,
                    format!("{context} slide master has an unsupported relationship"),
                );
            }
        }
        if relation_count != layout_ids.len() + 1 {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{context} slide master layout/theme list does not match relationships"),
            );
        }
        for id in layout_ids {
            let relationship = master.rels().get(&id).ok_or_else(|| {
                Error::Relationship(format!(
                    "{context} slide layout relationship '{id}' missing"
                ))
            })?;
            validate_internal(
                relationship,
                dialect,
                rt::SLIDE_LAYOUT,
                STRICT_LAYOUT_REL,
                "slide layout",
            )?;
            let layout_uri = relationship.target_partname()?;
            let layout = package.part(&layout_uri)?;
            if layout.content_type() != ct::PML_SLIDE_LAYOUT {
                return Err(Error::ContentType {
                    expected: ct::PML_SLIDE_LAYOUT.into(),
                    actual: layout.content_type().into(),
                });
            }
            if !layout_seen.insert(folded_ascii_name(
                layout_uri.as_str(),
                "source-backed registered slide-layout name",
            )?) {
                return refusal(
                    SlideCopyRefusal::AmbiguousTopology,
                    format!("{context} repeats a slide layout"),
                );
            }
            result.push(layout_uri);
        }
        let theme = master
            .rels()
            .iter()
            .find(|relationship| {
                relation_matches(relationship.reltype(), dialect, rt::THEME, STRICT_THEME_REL)
            })
            .ok_or_else(|| Error::Relationship(format!("{context} slide master theme missing")))?;
        validate_internal(theme, dialect, rt::THEME, STRICT_THEME_REL, "theme")?;
        let theme_part = package.part(&theme.target_partname()?)?;
        if theme_part.content_type() != ct::OFC_THEME || !theme_part.rels().is_empty() {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                format!("{context} theme is not a leaf part"),
            );
        }
    }
    Ok(result)
}

fn prove_graphs(
    source: &SourceBackedPackage,
    source_layout: &PackURI,
    destination: &SourceBackedPackage,
    destination_layout: &PackURI,
    dialect: Dialect,
) -> Result<[u8; 32]> {
    let source_graph = graph_digest(source, source_layout, dialect, "source")?;
    let destination_graph = graph_digest(destination, destination_layout, dialect, "destination")?;
    if source_graph != destination_graph {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            "layout/master/theme graphs are not equivalent",
        );
    }
    Ok(source_graph)
}

fn graph_digest(
    package: &SourceBackedPackage,
    layout_uri: &PackURI,
    dialect: Dialect,
    context: &'static str,
) -> Result<[u8; 32]> {
    let layout = package.part(layout_uri)?;
    let layout_data = layout.data()?;
    if validate_xml(
        layout_data.as_bytes(),
        b"sldLayout",
        RootNamespace::PresentationMl,
        true,
        true,
        false,
        false,
        false,
        "slide layout",
    )? != dialect
    {
        return refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            format!("{context} layout mixes dialects"),
        );
    }
    let layout_rel = single_relationship(
        layout.rels(),
        dialect,
        rt::SLIDE_MASTER,
        STRICT_MASTER_REL,
        "layout master",
    )?;
    let master_uri = layout_rel.target_partname()?;
    let master = package.part(&master_uri)?;
    let master_data = master.data()?;
    if master.content_type() != ct::PML_SLIDE_MASTER
        || validate_xml(
            master_data.as_bytes(),
            b"sldMaster",
            RootNamespace::PresentationMl,
            true,
            true,
            true,
            false,
            false,
            "slide master",
        )? != dialect
    {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("{context} master is incompatible"),
        );
    }
    let layout_ids = relationship_list(master_data.as_bytes(), b"sldLayoutIdLst", b"sldLayoutId")?;
    if !layout_ids.iter().any(|id| {
        master
            .rels()
            .get(id)
            .and_then(|r| r.target_partname().ok())
            .is_some_and(|uri| uri.is_equivalent_to(layout_uri))
    }) {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            format!("{context} selected layout is not registered"),
        );
    }
    let theme_rel = single_theme_relationship(master.rels(), dialect)?;
    let theme_uri = theme_rel.target_partname()?;
    let theme = package.part(&theme_uri)?;
    let theme_data = theme.data()?;
    if theme.content_type() != ct::OFC_THEME
        || validate_xml(
            theme_data.as_bytes(),
            b"theme",
            RootNamespace::DrawingMl,
            false,
            true,
            false,
            false,
            false,
            "theme",
        )? != dialect
        || !theme.rels().is_empty()
    {
        return refusal(
            SlideCopyRefusal::SharedOwner,
            format!("{context} theme is incompatible"),
        );
    }
    let mut digest = Sha256::new();
    for (name, view, data) in [
        ("layout", layout, layout_data),
        ("master", master, master_data),
        ("theme", theme, theme_data),
    ] {
        digest.update(name.as_bytes());
        digest.update(view.content_type().as_bytes());
        digest.update((data.as_bytes().len() as u64).to_le_bytes());
        digest.update(data.as_bytes());
        let mut rels = Vec::new();
        rels.try_reserve_exact(view.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed graph relationships",
                source,
            })?;
        rels.extend(view.rels().iter());
        rels.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
        for relationship in rels {
            if relationship.target_mode() != TargetMode::Internal
                || relationship.target_query().is_some()
                || relationship.target_fragment().is_some()
            {
                return refusal(
                    SlideCopyRefusal::UnsupportedRelationship,
                    format!("{context} graph has a non-exact edge"),
                );
            }
            digest.update(relationship.r_id().as_bytes());
            digest.update(relationship.reltype().as_bytes());
            digest.update(relationship.target_ref().as_bytes());
        }
    }
    Ok(digest.finalize().into())
}

fn single_relationship<'a>(
    relationships: &'a Relationships,
    dialect: Dialect,
    transitional: &str,
    strict: &str,
    label: &'static str,
) -> Result<&'a litchi_opc::Relationship> {
    let mut found = None;
    for relationship in relationships.iter() {
        if !relation_matches(relationship.reltype(), dialect, transitional, strict) {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                format!("{label} has an unsupported edge"),
            );
        }
        if found.is_some() {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{label} has multiple edges"),
            );
        }
        validate_internal(relationship, dialect, transitional, strict, label)?;
        found = Some(relationship);
    }
    found.ok_or_else(|| Error::Relationship(format!("{label} is missing")))
}

fn single_theme_relationship(
    relationships: &Relationships,
    dialect: Dialect,
) -> Result<&litchi_opc::Relationship> {
    let mut found = None;
    for relationship in relationships.iter() {
        if relation_matches(relationship.reltype(), dialect, rt::THEME, STRICT_THEME_REL) {
            if found.is_some() {
                return refusal(
                    SlideCopyRefusal::AmbiguousTopology,
                    "master has multiple theme edges",
                );
            }
            validate_internal(
                relationship,
                dialect,
                rt::THEME,
                STRICT_THEME_REL,
                "master theme",
            )?;
            found = Some(relationship);
        } else if relation_matches(
            relationship.reltype(),
            dialect,
            rt::SLIDE_LAYOUT,
            STRICT_LAYOUT_REL,
        ) {
            validate_internal(
                relationship,
                dialect,
                rt::SLIDE_LAYOUT,
                STRICT_LAYOUT_REL,
                "master slide layout",
            )?;
        } else {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                "master theme has an unsupported edge",
            );
        }
    }
    found.ok_or_else(|| Error::Relationship("master theme is missing".into()))
}

fn validate_internal(
    relationship: &litchi_opc::Relationship,
    dialect: Dialect,
    transitional: &str,
    strict: &str,
    label: &'static str,
) -> Result<()> {
    if !relation_matches(relationship.reltype(), dialect, transitional, strict)
        || relationship.is_external()
        || relationship.target_mode() != TargetMode::Internal
        || relationship.target_query().is_some()
        || relationship.target_fragment().is_some()
    {
        return refusal(
            SlideCopyRefusal::UnsupportedRelationship,
            format!("{label} is not an exact internal edge"),
        );
    }
    Ok(())
}

fn relation_matches(actual: &str, dialect: Dialect, transitional: &str, strict: &str) -> bool {
    actual
        == match dialect {
            Dialect::Transitional => transitional,
            Dialect::Strict => strict,
        }
}

fn validate_presentation_relationships(
    relationships: &Relationships,
    context: &'static str,
) -> Result<()> {
    for relationship in relationships.iter() {
        if relationship.is_external()
            || relationship.target_mode() != TargetMode::Internal
            || relationship.target_query().is_some()
            || relationship.target_fragment().is_some()
        {
            return refusal(
                SlideCopyRefusal::UnsupportedRelationship,
                format!("{context} presentation relationships are not exact internal edges"),
            );
        }
    }
    Ok(())
}

fn validate_slide_refs(references: &[SlideReference], context: &'static str) -> Result<()> {
    let mut ids = HashSet::new();
    let mut relationships = HashSet::new();
    ids.try_reserve(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide ID proof",
            source,
        })?;
    relationships
        .try_reserve(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide relationship proof",
            source,
        })?;
    for reference in references {
        if !(MIN_SLIDE_ID..=MAX_SLIDE_ID).contains(&reference.id())
            || !ids.insert(reference.id())
            || !relationships.insert(reference.relationship_id())
        {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{context} has a bad or duplicate slide identity"),
            );
        }
    }
    Ok(())
}

fn validate_registered_slide_parts(
    package: &SourceBackedPackage,
    slides: &[Arc<super::source::SourceSlideData>],
    context: &'static str,
) -> Result<()> {
    let mut expected = HashSet::new();
    expected
        .try_reserve(slides.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed registered slide-part index",
            source,
        })?;
    for slide in slides {
        if !expected.insert(folded_ascii_name(
            slide.part_uri.as_str(),
            "source-backed registered slide-part name",
        )?) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{context} repeats a registered slide Part"),
            );
        }
    }
    let mut registered = 0usize;
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE {
            continue;
        }
        registered = registered
            .checked_add(1)
            .ok_or_else(|| invalid("registered slide count overflow"))?;
        let key = folded_ascii_name(
            part.partname().as_str(),
            "source-backed package slide-part name",
        )?;
        if !expected.remove(&key) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                format!("{context} contains an unregistered slide Part"),
            );
        }
    }
    if registered != slides.len() || !expected.is_empty() {
        return refusal(
            SlideCopyRefusal::AmbiguousTopology,
            format!("{context} slide registry does not cover the package"),
        );
    }
    Ok(())
}

fn allocate_slide_id(references: &[SlideReference]) -> Result<u32> {
    let mut used = Vec::new();
    used.try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed used slide IDs",
            source,
        })?;
    used.extend(references.iter().map(SlideReference::id));
    used.sort_unstable();
    let mut candidate = MIN_SLIDE_ID;
    for used_id in used {
        if used_id == candidate {
            candidate = candidate
                .checked_add(1)
                .ok_or_else(|| Error::SlideCopyPlan {
                    kind: SlideCopyRefusal::AmbiguousTopology,
                    detail: "slide ID space is exhausted".into(),
                })?;
        } else if used_id > candidate {
            break;
        }
    }
    if candidate <= MAX_SLIDE_ID {
        Ok(candidate)
    } else {
        refusal(
            SlideCopyRefusal::AmbiguousTopology,
            "slide ID space is exhausted",
        )
    }
}

fn allocate_relationship_id(relationships: &Relationships) -> Result<String> {
    let maximum = relationships
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("relationship candidate count overflow"))?;
    for candidate in 1..=maximum {
        let id = format!("rId{candidate}");
        if relationships.get(&id).is_none() {
            return Ok(id);
        }
    }
    refusal(
        SlideCopyRefusal::AmbiguousTopology,
        "relationship ID space is exhausted",
    )
}

fn allocate_slide_uri(package: &SourceBackedPackage) -> Result<PackURI> {
    let physical_count = package.physical_member_names().len();
    let mut physical_names = HashSet::new();
    physical_names
        .try_reserve(physical_count)
        .map_err(|source| Error::Allocation {
            resource: "source-backed physical member-name index",
            source,
        })?;
    for name in package.physical_member_names() {
        let key = folded_ascii_name(name, "source-backed physical member name")?;
        if !physical_names.insert(key) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "destination contains equivalent physical member names",
            );
        }
    }
    let maximum = physical_count
        .checked_add(1)
        .ok_or_else(|| invalid("slide part-name candidate count overflow"))?;
    for index in 1..=maximum {
        let uri = PackURI::new(format!("/ppt/slides/slide{index}.xml"))
            .map_err(|error| Error::Uri(error.to_string()))?;
        let relationship_uri = uri
            .rels_uri()
            .map_err(|error| Error::Uri(error.to_string()))?;
        let member_key = folded_ascii_name(
            uri.membername(),
            "source-backed candidate slide member name",
        )?;
        let relationship_key = folded_ascii_name(
            relationship_uri.membername(),
            "source-backed candidate slide relationship member name",
        )?;
        if physical_names.contains(&member_key) || physical_names.contains(&relationship_key) {
            continue;
        }
        return Ok(uri);
    }
    refusal(
        SlideCopyRefusal::AmbiguousTopology,
        "slide part-name space is exhausted",
    )
}

fn allocate_media_uri(package: &SourceBackedPackage, source_uri: &PackURI) -> Result<PackURI> {
    let physical_count = package.physical_member_names().len();
    let mut physical_names = HashSet::new();
    physical_names
        .try_reserve(physical_count)
        .map_err(|source| Error::Allocation {
            resource: "source-backed physical member-name index",
            source,
        })?;
    for name in package.physical_member_names() {
        let key = folded_ascii_name(name, "source-backed physical member name")?;
        if !physical_names.insert(key) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "destination contains equivalent physical member names",
            );
        }
    }
    let source_member = source_uri.membername();
    let source_name = source_member
        .rsplit('/')
        .next()
        .filter(|name| !name.is_empty())
        .ok_or_else(|| Error::SlideCopyPlan {
            kind: SlideCopyRefusal::AmbiguousTopology,
            detail: "embedded image target has no media member name".into(),
        })?;
    let (stem, extension) = source_name.rsplit_once('.').unwrap_or((source_name, ""));
    let maximum = physical_count
        .checked_add(1)
        .ok_or_else(|| invalid("media part-name candidate count overflow"))?;
    for index in 1..=maximum {
        let name = if extension.is_empty() {
            format!("{stem}-copy{index}")
        } else {
            format!("{stem}-copy{index}.{extension}")
        };
        let uri = PackURI::new(format!("/ppt/media/{name}"))
            .map_err(|error| Error::Uri(error.to_string()))?;
        let relationship_uri = uri
            .rels_uri()
            .map_err(|error| Error::Uri(error.to_string()))?;
        let member_key = folded_ascii_name(
            uri.membername(),
            "source-backed candidate image member name",
        )?;
        let relationship_key = folded_ascii_name(
            relationship_uri.membername(),
            "source-backed candidate image relationship member name",
        )?;
        if physical_names.contains(&member_key) || physical_names.contains(&relationship_key) {
            continue;
        }
        return Ok(uri);
    }
    refusal(
        SlideCopyRefusal::AmbiguousTopology,
        "media part-name space is exhausted",
    )
}

fn folded_ascii_name(value: &str, resource: &'static str) -> Result<String> {
    let mut folded = String::new();
    folded
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    folded.push_str(value);
    folded.make_ascii_lowercase();
    Ok(folded)
}

fn destination_names(
    package: &SourceBackedPackage,
    presentation: &PartView<'_>,
    references: &[SlideReference],
) -> Result<HashSet<String>> {
    let mut names = HashSet::new();
    names
        .try_reserve(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed destination slide names",
            source,
        })?;
    for reference in references {
        let relationship = presentation
            .rels()
            .get(reference.relationship_id())
            .ok_or_else(|| Error::Relationship("destination slide relationship missing".into()))?;
        let uri = relationship.target_partname()?;
        let view = package.part(&uri)?;
        if view.content_type() != ct::PML_SLIDE {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE.into(),
                actual: view.content_type().into(),
            });
        }
        let data = view.data()?;
        let name = crate::namespace::presentation_name(data.as_bytes())?
            .filter(|name| !name.is_empty())
            .ok_or_else(|| Error::SlideCopyPlan {
                kind: SlideCopyRefusal::AmbiguousTopology,
                detail: "destination slide name is missing or empty".into(),
            })?;
        if !names.insert(normalize_name(&name)?) {
            return refusal(
                SlideCopyRefusal::AmbiguousTopology,
                "destination contains duplicate slide names",
            );
        }
    }
    Ok(names)
}

fn normalize_name(name: &str) -> Result<String> {
    let bytes = name
        .chars()
        .try_fold(0usize, |total, character| {
            character.to_lowercase().try_fold(total, |total, lowered| {
                total.checked_add(lowered.len_utf8())
            })
        })
        .ok_or_else(|| invalid("normalized slide name length overflow"))?;
    let mut normalized = String::new();
    normalized
        .try_reserve_exact(bytes)
        .map_err(|source| Error::Allocation {
            resource: "source-backed normalized slide name",
            source,
        })?;
    normalized.extend(name.chars().flat_map(char::to_lowercase));
    Ok(normalized)
}

fn reject_package_features(
    package: &SourceBackedPackage,
    presentation_xml: &[u8],
    context: &'static str,
) -> Result<()> {
    if package.has_encrypted_entries() {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            format!("{context} contains encrypted ZIP members"),
        );
    }
    if !package.non_part_members().is_empty() {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            format!("{context} contains unknown non-Part members"),
        );
    }
    if package.rels().iter().any(is_signature_relationship)
        || package.iter_parts().any(|part| {
            part.partname()
                .as_str()
                .to_ascii_lowercase()
                .starts_with("/_xmlsignatures/")
                || part.content_type().contains("digital-signature")
                || part.rels().iter().any(is_signature_relationship)
        })
    {
        return refusal(
            SlideCopyRefusal::SignedPackage,
            format!("{context} contains signature infrastructure"),
        );
    }
    if package.rels().iter().any(is_macro_relationship)
        || package.iter_parts().any(|part| {
            matches!(
                part.content_type(),
                ct::OFC_VBA_PROJECT
                    | ct::OFC_VBA_PROJECT_SIGNATURE
                    | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
                    | ct::PML_PRES_MACRO_MAIN
                    | ct::PML_SLIDESHOW_MACRO_MAIN
                    | ct::PML_TEMPLATE_MACRO_MAIN
            ) || part.rels().iter().any(is_macro_relationship)
        })
    {
        return refusal(
            SlideCopyRefusal::UnknownPhysicalMember,
            format!("{context} contains macro/VBA infrastructure"),
        );
    }
    let text =
        std::str::from_utf8(presentation_xml).map_err(|error| Error::Xml(error.to_string()))?;
    if crate::presentation_properties::metadata::protection::Settings::parse_xml(text)?
        .is_protected()
    {
        return refusal(
            SlideCopyRefusal::ProtectedPresentation,
            format!("{context} is actively protected"),
        );
    }
    Ok(())
}

fn is_signature_relationship(relationship: &litchi_opc::Relationship) -> bool {
    matches!(
        relationship.reltype(),
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_macro_relationship(relationship: &litchi_opc::Relationship) -> bool {
    matches!(
        relationship.reltype(),
        rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
    )
}

fn relationship_list(xml: &[u8], list_name: &[u8], entry_name: &[u8]) -> Result<Vec<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut list_depth = None;
    let mut lists = 0usize;
    let mut values = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml = is_pml(&namespace);
        match event {
            Event::Start(element) => {
                if pml && depth == 1 && element.local_name().as_ref() == list_name {
                    lists += 1;
                    list_depth = Some(depth + 1);
                } else if list_depth == Some(depth)
                    && pml
                    && element.local_name().as_ref() == entry_name
                {
                    return Err(Error::Invalid(
                        "relationship-list entries must be empty".into(),
                    ));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
                if pml && child_depth == 2 && element.local_name().as_ref() == list_name {
                    lists += 1;
                } else if list_depth == Some(depth)
                    && pml
                    && element.local_name().as_ref() == entry_name
                {
                    let id = crate::namespace::relationship_attribute_value(
                        &element,
                        b"id",
                        reader.decoder(),
                        reader.resolver(),
                    )?
                    .ok_or_else(|| invalid("relationship-list entry lacks r:id"))?;
                    values.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "source-backed relationship list",
                        source,
                    })?;
                    values.push(id);
                }
            },
            Event::End(element) => {
                if pml && list_depth == Some(depth) && element.local_name().as_ref() == list_name {
                    list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("XML depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::MarkupCompatibility,
                    "relationship list contains DTD or PI",
                );
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if lists != 1 || depth != 0 || list_depth.is_some() {
        return Err(Error::Invalid("relationship list is malformed".into()));
    }
    Ok(values)
}

fn validate_xml(
    xml: &[u8],
    root: &[u8],
    root_namespace: RootNamespace,
    allow_pml: bool,
    allow_dml: bool,
    allow_presentation_relationship_attributes: bool,
    reject_dependency_surfaces: bool,
    allow_unknown_namespaces: bool,
    context: &'static str,
) -> Result<Dialect> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_content_seen = false;
    let mut transitional = false;
    let mut strict = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if root_closed {
                    return Err(Error::Xml(format!(
                        "{context} contains more than one root element"
                    )));
                }
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML node count overflow"))?;
                if nodes > MAX_XML_NODES || depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "source-backed cross-copy XML",
                        limit: MAX_XML_NODES,
                    });
                }
                let value = match &namespace {
                    ResolveResult::Bound(Namespace(value)) => *value,
                    ResolveResult::Unknown(_) if allow_unknown_namespaces => {
                        if is_start {
                            depth += 1;
                        }
                        continue;
                    },
                    _ => {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            format!("{context} uses an unresolved namespace"),
                        );
                    },
                };
                if value == MCE_NAMESPACE || value == STRICT_MCE_NAMESPACE {
                    return refusal(
                        SlideCopyRefusal::MarkupCompatibility,
                        format!("{context} contains MCE"),
                    );
                }
                let pml = value == crate::namespace::PRESENTATIONML_NAMESPACE
                    || value == STRICT_PML_NAMESPACE;
                let dml = value == TRANSITIONAL_DML_NAMESPACE || value == STRICT_DML_NAMESPACE;
                let pml_is_transitional = value == crate::namespace::PRESENTATIONML_NAMESPACE;
                let pml_is_strict = value == STRICT_PML_NAMESPACE;
                let creation_id =
                    value == P14_NAMESPACE && element.local_name().as_ref() == b"creationId";
                if (!pml || !allow_pml)
                    && (!dml || !allow_dml)
                    && !creation_id
                    && !allow_unknown_namespaces
                {
                    return refusal(
                        SlideCopyRefusal::UnknownSemanticSurface,
                        format!(
                            "{context} contains unsupported namespace '{}' on element '{}'",
                            String::from_utf8_lossy(value),
                            String::from_utf8_lossy(element.local_name().as_ref())
                        ),
                    );
                }
                if pml {
                    if value == crate::namespace::PRESENTATIONML_NAMESPACE {
                        transitional = true;
                    } else {
                        strict = true;
                    }
                }
                if dml {
                    if value == TRANSITIONAL_DML_NAMESPACE {
                        transitional = true;
                    } else {
                        strict = true;
                    }
                }
                if creation_id {
                    if is_start {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            format!("{context} contains a non-empty p14:creationId"),
                        );
                    }
                    let mut value_seen = false;
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let key = attribute.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            continue;
                        }
                        if key != b"val"
                            || value_seen
                            || attribute.value.is_empty()
                            || !attribute.value.iter().all(u8::is_ascii_digit)
                            || std::str::from_utf8(attribute.value.as_ref())
                                .ok()
                                .and_then(|value| value.parse::<u32>().ok())
                                .is_none()
                        {
                            return refusal(
                                SlideCopyRefusal::UnknownSemanticSurface,
                                format!("{context} contains a non-canonical p14:creationId"),
                            );
                        }
                        value_seen = true;
                    }
                    if !value_seen {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            format!("{context} contains p14:creationId without val"),
                        );
                    }
                }
                if !root_seen {
                    root_seen = true;
                    if element.local_name().as_ref() != root || !root_namespace.matches(pml, dml) {
                        return refusal(
                            SlideCopyRefusal::UnknownSemanticSurface,
                            format!("{context} has an unexpected root"),
                        );
                    }
                    if !is_start {
                        root_closed = true;
                    }
                } else if reject_dependency_surfaces
                    && unsupported_surface(element.local_name().as_ref())
                {
                    return refusal(
                        SlideCopyRefusal::UnknownSemanticSurface,
                        format!("{context} contains an unsupported dependency surface"),
                    );
                }
                let allow_source_slide_image_relationship_attributes =
                    matches!(context, "source slide" | "staged source slide")
                        && matches!(root_namespace, RootNamespace::PresentationMl)
                        && dml
                        && element.local_name().as_ref() == b"blip";
                let expected_relationship_namespace = if pml_is_strict {
                    Some(STRICT_REL_NAMESPACE)
                } else if pml_is_transitional {
                    Some(TRANSITIONAL_REL_NAMESPACE)
                } else if allow_source_slide_image_relationship_attributes
                    && strict
                    && !transitional
                {
                    Some(STRICT_REL_NAMESPACE)
                } else if allow_source_slide_image_relationship_attributes
                    && transitional
                    && !strict
                {
                    Some(TRANSITIONAL_REL_NAMESPACE)
                } else {
                    None
                };
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let key = attribute.key.as_ref();
                    let attribute_namespace = reader.resolver().resolve_attribute(attribute.key).0;
                    let relationship_namespace = match attribute_namespace {
                        ResolveResult::Bound(Namespace(value))
                            if value == TRANSITIONAL_REL_NAMESPACE
                                || value == STRICT_REL_NAMESPACE =>
                        {
                            Some(value)
                        },
                        _ => None,
                    };
                    if key.starts_with(b"r:") || relationship_namespace.is_some() {
                        let allowed_presentation_attribute =
                            allow_presentation_relationship_attributes
                                && pml
                                && attribute.key.local_name().as_ref() == b"id"
                                && relationship_namespace == expected_relationship_namespace
                                && matches!(
                                    element.local_name().as_ref(),
                                    b"sldId"
                                        | b"sldMasterId"
                                        | b"sldLayoutId"
                                        | b"notesMasterId"
                                        | b"handoutMasterId"
                                );
                        let allowed_image_attribute =
                            allow_source_slide_image_relationship_attributes
                                && attribute.key.local_name().as_ref() == b"embed"
                                && relationship_namespace == expected_relationship_namespace;
                        let allowed = allowed_presentation_attribute || allowed_image_attribute;
                        if !allowed {
                            return refusal(
                                SlideCopyRefusal::UnsupportedRelationship,
                                format!("{context} contains a relationship-qualified attribute"),
                            );
                        }
                    }
                    if matches!(attribute_namespace, ResolveResult::Bound(Namespace(value)) if value == MCE_NAMESPACE || value == STRICT_MCE_NAMESPACE)
                    {
                        return refusal(
                            SlideCopyRefusal::MarkupCompatibility,
                            format!("{context} contains an MCE attribute"),
                        );
                    }
                }
                if is_start {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("XML depth underflow"))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::MarkupCompatibility,
                    format!("{context} contains DTD or PI"),
                );
            },
            Event::Decl(_) if root_seen || declaration_seen || prolog_content_seen => {
                return Err(Error::Xml(format!(
                    "{context} contains a misplaced or repeated XML declaration"
                )));
            },
            Event::Decl(_) => declaration_seen = true,
            Event::Text(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::Xml(format!(
                        "{context} contains text outside its root"
                    )));
                }
                if !root_seen && !text.as_ref().is_empty() {
                    prolog_content_seen = true;
                }
            },
            Event::Comment(_) if !root_seen => prolog_content_seen = true,
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Xml(format!(
                    "{context} contains character data outside its root"
                )));
            },
            Event::GeneralRef(reference) if !valid_xml_reference(&reference) => {
                return Err(Error::Xml(format!(
                    "{context} contains an invalid XML character reference"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(Error::Xml(format!("{context} is not one closed root")));
    }
    match (transitional, strict) {
        (true, false) => Ok(Dialect::Transitional),
        (false, true) => Ok(Dialect::Strict),
        (true, true) => refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            format!("{context} mixes dialects"),
        ),
        (false, false) => refusal(
            SlideCopyRefusal::UnknownSemanticSurface,
            format!("{context} has no OOXML namespace"),
        ),
    }
}

fn valid_xml_reference(reference: &BytesRef<'_>) -> bool {
    let bytes: &[u8] = reference;
    matches!(bytes, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
        || (reference.is_char_ref()
            && reference
                .resolve_char_ref()
                .ok()
                .flatten()
                .is_some_and(|value| {
                    matches!(value, '\u{9}' | '\u{a}' | '\u{d}')
                        || ('\u{20}'..='\u{d7ff}').contains(&value)
                        || ('\u{e000}'..='\u{fffd}').contains(&value)
                        || ('\u{10000}'..='\u{10ffff}').contains(&value)
                }))
}

fn unsupported_surface(local: &[u8]) -> bool {
    matches!(
        local,
        b"tbl"
            | b"graphicFrame"
            | b"contentPart"
            | b"oleObj"
            | b"olePic"
            | b"audio"
            | b"video"
            | b"media"
            | b"extLst"
            | b"timing"
            | b"custDataLst"
            | b"embeddedFontLst"
            | b"notes"
            | b"comment"
            | b"AlternateContent"
            | b"Choice"
            | b"Fallback"
    )
}

fn is_pml(namespace: &ResolveResult<'_>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            let value: &[u8] = value;
            value == crate::namespace::PRESENTATIONML_NAMESPACE || value == STRICT_PML_NAMESPACE
        },
        _ => false,
    }
}

fn is_dml(namespace: &ResolveResult<'_>) -> bool {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            let value: &[u8] = value;
            value == TRANSITIONAL_DML_NAMESPACE || value == STRICT_DML_NAMESPACE
        },
        _ => false,
    }
}

#[derive(Clone)]
struct SlideElement {
    id: u32,
    relationship_id: String,
}

fn slide_elements(xml: &[u8]) -> Result<Vec<SlideElement>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut list_depth = None;
    let mut lists = 0usize;
    let mut elements = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml = is_pml(&namespace);
        let event = event.into_owned();
        drop(namespace);
        match event {
            Event::Start(element) => {
                if pml && depth == 1 && element.local_name().as_ref() == b"sldIdLst" {
                    lists = lists
                        .checked_add(1)
                        .ok_or_else(|| invalid("destination slide-list count overflow"))?;
                    if lists != 1 {
                        return refusal(
                            SlideCopyRefusal::AmbiguousTopology,
                            "presentation contains multiple slide ID lists",
                        );
                    }
                    list_depth = Some(depth + 1);
                } else if list_depth == Some(depth)
                    && pml
                    && element.local_name().as_ref() == b"sldId"
                {
                    return Err(invalid("destination slide IDs must be empty"));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if pml && depth == 1 && element.local_name().as_ref() == b"sldIdLst" {
                    return refusal(
                        SlideCopyRefusal::AmbiguousTopology,
                        "presentation slide ID list cannot be empty",
                    );
                }
                if list_depth == Some(depth) && pml && element.local_name().as_ref() == b"sldId" {
                    let id = litchi_ooxml_common::xml::unqualified_attribute_value(
                        &element,
                        b"id",
                        reader.decoder(),
                    )?
                    .ok_or_else(|| invalid("slide ID lacks id"))?
                    .parse::<u32>()
                    .map_err(|_| invalid("slide ID is invalid"))?;
                    let relationship_id = crate::namespace::relationship_attribute_value(
                        &element,
                        b"id",
                        reader.decoder(),
                        reader.resolver(),
                    )?
                    .ok_or_else(|| invalid("slide ID lacks r:id"))?;
                    elements
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "source-backed staged slide bindings",
                            source,
                        })?;
                    elements.push(SlideElement {
                        id,
                        relationship_id,
                    });
                }
            },
            Event::End(element) => {
                if pml && list_depth == Some(depth) && element.local_name().as_ref() == b"sldIdLst"
                {
                    list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("slide XML depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return refusal(
                    SlideCopyRefusal::MarkupCompatibility,
                    "presentation contains DTD or PI",
                );
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || list_depth.is_some() || lists != 1 {
        return Err(invalid("presentation XML is unterminated"));
    }
    Ok(elements)
}

fn validate_candidate_bindings<'a>(
    xml: &[u8],
    current: impl ExactSizeIterator<Item = (u32, &'a str)>,
    position: usize,
    id: u32,
    relationship_id: &str,
) -> Result<()> {
    let actual = slide_elements(xml)?;
    let expected_len = current
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("staged slide binding count overflow"))?;
    if actual.len() != expected_len || position > current.len() {
        return Err(Error::Invalid(
            "staged presentation slide bindings differ from plan".into(),
        ));
    }
    let mut current = current;
    for (index, actual) in actual.iter().enumerate() {
        let matches = if index == position {
            actual.id == id && actual.relationship_id == relationship_id
        } else {
            let Some((expected_id, expected_relationship_id)) = current.next() else {
                return Err(Error::Invalid(
                    "staged presentation slide bindings differ from plan".into(),
                ));
            };
            actual.id == expected_id && actual.relationship_id == expected_relationship_id
        };
        if !matches {
            return Err(Error::Invalid(
                "staged presentation slide bindings differ from plan".into(),
            ));
        }
    }
    if current.next().is_some() {
        return Err(Error::Invalid(
            "staged presentation slide bindings differ from plan".into(),
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn digest_touched(
    source_presentation: &[u8],
    destination_presentation: &[u8],
    source_slide: &[u8],
    target_presentation: &[u8],
    graph: [u8; 32],
    source_slide_uri: &PackURI,
    source_layout_uri: &PackURI,
    destination_layout_uri: &PackURI,
    target_slide_uri: &PackURI,
    slide_id: u32,
    relationship_id: &str,
    source_name: &str,
    destination_names: &HashSet<String>,
    image: Option<&PreparedImage>,
) -> Result<[u8; 32]> {
    let mut digest = Sha256::new();
    for bytes in [
        source_presentation,
        destination_presentation,
        source_slide,
        target_presentation,
    ] {
        digest.update((bytes.len() as u64).to_le_bytes());
        digest.update(bytes);
    }
    digest.update(graph);
    for uri in [
        source_slide_uri,
        source_layout_uri,
        destination_layout_uri,
        target_slide_uri,
    ] {
        digest.update(uri.as_str().as_bytes());
        digest.update([0]);
    }
    digest.update(slide_id.to_le_bytes());
    digest.update(relationship_id.as_bytes());
    digest.update(source_name.as_bytes());
    if let Some(image) = image {
        for value in [
            image.source_uri.as_str(),
            image.target_uri.as_str(),
            image.relationship_id.as_str(),
            image.relationship_type.as_str(),
            image.content_type.as_str(),
        ] {
            digest.update(value.as_bytes());
            digest.update([0]);
        }
        digest.update((image.bytes.len() as u64).to_le_bytes());
        digest.update(&image.bytes);
    }
    let mut names = Vec::new();
    names
        .try_reserve_exact(destination_names.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed touched-name digest",
            source,
        })?;
    names.extend(destination_names.iter());
    names.sort_unstable();
    for name in names {
        digest.update(name.as_bytes());
        digest.update([0]);
    }
    Ok(digest.finalize().into())
}

fn reserve_memory(
    context: Option<&ExecutionContext>,
    bytes: usize,
    limits: litchi_opc::ReadLimits,
) -> Result<Option<Arc<Reservation>>> {
    let bytes = bytes
        .checked_add(MAX_STAGING_OVERHEAD)
        .ok_or_else(|| invalid("cross-copy staging size overflow"))?;
    let bytes_u64 = u64::try_from(bytes).map_err(|_| invalid("cross-copy staging exceeds u64"))?;
    if bytes_u64 > limits.max_total_part_bytes() {
        return Err(Error::Limit {
            resource: "source-backed cross-copy staged bytes",
            limit: usize::try_from(limits.max_total_part_bytes()).unwrap_or(usize::MAX),
        });
    }
    let Some(context) = context else {
        return Ok(None);
    };
    context
        .reserve(Resource::Memory, bytes_u64)
        .map(Arc::new)
        .map(Some)
        .map_err(map_execution_error)
}

fn clone_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(bytes);
    Ok(output)
}
fn map_execution_error(error: ExecutionError) -> Error {
    Error::Opc(match error {
        ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
        other => litchi_opc::OpcError::Execution(other),
    })
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn refusal<T>(kind: SlideCopyRefusal, detail: impl Into<String>) -> Result<T> {
    Err(Error::SlideCopyPlan {
        kind,
        detail: detail.into(),
    })
}

struct SourceCheckedWriter<W> {
    inner: W,
    source: SourceBackedPresentation,
    state: Arc<Mutex<SourceCheckState>>,
}

#[derive(Default)]
struct SourceCheckState {
    accepted: u64,
    source_failure: Option<litchi_opc::OpcError>,
}

impl<W: Write> Write for SourceCheckedWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.check_source()?;
        let written = self.inner.write(bytes)?;
        if written <= bytes.len() {
            let mut state = lock_source_state(&self.state);
            state.accepted = state
                .accepted
                .checked_add(u64::try_from(written).unwrap_or(u64::MAX))
                .unwrap_or(u64::MAX);
        }
        self.check_source()?;
        Ok(written)
    }
    fn flush(&mut self) -> io::Result<()> {
        self.check_source()?;
        self.inner.flush()?;
        self.check_source()
    }
}

impl<W> SourceCheckedWriter<W> {
    fn check_source(&self) -> io::Result<()> {
        match self.source.check_source() {
            Ok(()) => Ok(()),
            Err(Error::Opc(error)) => {
                lock_source_state(&self.state).source_failure = Some(error);
                Err(io::Error::other(
                    "PPTX source check failed during publication",
                ))
            },
            Err(error) => Err(io::Error::other(error)),
        }
    }
}

fn lock_source_state(
    state: &Mutex<SourceCheckState>,
) -> std::sync::MutexGuard<'_, SourceCheckState> {
    state
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

fn accepted_bytes(state: &Mutex<SourceCheckState>) -> u64 {
    lock_source_state(state).accepted
}

fn take_source_failure(state: &Mutex<SourceCheckState>) -> Option<litchi_opc::OpcError> {
    lock_source_state(state).source_failure.take()
}

fn source_failure_with_progress(written: u64, source: litchi_opc::OpcError) -> Error {
    if written == 0 {
        Error::Opc(source)
    } else {
        Error::Opc(litchi_opc::OpcError::IncompleteOutput {
            written,
            source: Box::new(source),
        })
    }
}
