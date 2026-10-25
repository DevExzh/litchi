//! Package-aware custom-show and section mutation for `PowerPoint` presentations.
//!
//! Custom shows reference slide relationships while `PowerPoint` 2010 sections
//! reference stable numeric presentation slide IDs. This module resolves both
//! representations into one validated graph and patches only the corresponding
//! children of `ppt/presentation.xml`.

use super::super::custom_show::{List as ShowList, Show};
use super::super::sections::{List as SectionList, Section};
use super::super::{BoundedString, Limits, checked_output_len, process_ooxml};
use super::model::{Graph, Reference};
use crate::package::signature::settle_unsigned_change;
use crate::presentation_properties::metadata::new_guid;
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::OpcPackage;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use std::collections::{HashMap, HashSet};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SLIDE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const MAX_SLIDE_ID: u32 = 2_147_483_647;

/// Load and validate the presentation structure graph.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, limits: Limits) -> Result<Graph> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    parse_structure_blob(package, presentation.blob(), true, limits)
}

/// Atomically replace custom shows and sections while preserving unrelated XML.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub(crate) fn store(package: &mut OpcPackage, value: &Graph) -> Result<()> {
    let limits = Limits::from_source_limits(package.read_limits());
    validate_graph(package, value, limits)?;
    let presentation_name = package.main_document_part()?.partname().clone();
    let original = package.get_part(&presentation_name)?.blob().to_vec();
    let original_graph = parse_structure_blob(package, &original, true, limits)?;
    let staged = rewrite(&original, &original_graph, value, limits)?;

    if staged == original {
        return Ok(());
    }

    let reparsed = parse_structure_blob(package, &staged, true, limits)?;
    if reparsed.slides != value.slides
        || reparsed.custom_shows.shows != value.custom_shows.shows
        || reparsed.sections != value.sections
    {
        return Err(invalid("staged presentation structure did not round-trip"));
    }
    // This refuses a signed package before the mutation; an unsigned one
    // only has its OPC signature-policy state settled, here as before.
    settle_unsigned_change(package)?;
    package.get_part_mut(&presentation_name)?.set_blob(staged);
    Ok(())
}

/// Rewrite only the changed structure owners in one validated presentation
/// document. The slide relationship graph is intentionally immutable here;
/// slide insertion and removal belong to the presentation package facade.
pub(crate) fn rewrite(
    source: &[u8],
    original: &Graph,
    value: &Graph,
    limits: Limits,
) -> Result<Vec<u8>> {
    validate_detached_graph(original, value, limits)?;
    if equivalent_graph(original, value) {
        if source.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                source.len(),
                limits.max_output_bytes(),
            ));
        }
        return Ok(source.to_vec());
    }

    if source.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            source.len(),
            limits.max_output_bytes(),
        ));
    }

    let (p_namespace, _r_namespace) = document_namespaces(source);
    let mut staged = source.to_vec();
    if original.custom_shows.shows != value.custom_shows.shows {
        // Reuse the source-preserving custom-show wire layout.  Rebuilding the
        // whole custShowLst from the graph loses nested future attributes and
        // elements that the typed graph deliberately does not model.
        staged = super::super::custom_show::rewrite_source(
            &staged,
            &original.custom_shows,
            &value.custom_shows,
            limits,
        )?;
    }
    if original.sections != value.sections {
        let section_xml = write_section_extension(&value.sections, p_namespace, limits)?;
        staged = patch_sections(&staged, section_xml.as_bytes(), p_namespace, limits)?;
    }

    let reparsed = parse_detached(&staged, &value.slides, limits)?;
    if !equivalent_graph(&reparsed, value) {
        return Err(invalid("staged presentation structure did not round-trip"));
    }
    Ok(staged)
}

/// Parse a rewritten presentation against its already validated relationship
/// topology without needing to reopen any slide part.
pub(crate) fn parse_detached(xml: &[u8], slides: &[Reference], limits: Limits) -> Result<Graph> {
    if xml.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            xml.len(),
            limits.max_input_bytes(),
        ));
    }
    let processed = process_ooxml(xml, limits)?;
    if processed.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            processed.len(),
            limits.max_output_bytes(),
        ));
    }
    let (raw_slides, raw_shows) = parse_core(processed.as_ref(), limits)?;
    let expected_slides = slides
        .iter()
        .map(|slide| (slide.slide_id, slide.relationship_id.clone()))
        .collect::<Vec<_>>();
    if raw_slides != expected_slides {
        return Err(invalid(
            "presentation slide relationship topology changed in a structure transaction",
        ));
    }
    let custom_shows = resolve_custom_shows(raw_shows, slides, true)?;
    let graph = Graph {
        slides: slides.to_vec(),
        custom_shows,
        sections: SectionList::from_xml(xml, limits)?,
    };
    validate_graph_shape(&graph, limits)?;
    Ok(graph)
}

/// Validate detached edits against the source's immutable slide graph.
pub(crate) fn validate_detached_graph(
    original: &Graph,
    value: &Graph,
    limits: Limits,
) -> Result<()> {
    if value.slides != original.slides {
        return Err(invalid(
            "presentation slide ordering and relationship topology are immutable in a structure transaction",
        ));
    }
    validate_graph_shape(value, limits)
}

pub(crate) fn equivalent_graph(left: &Graph, right: &Graph) -> bool {
    left.slides == right.slides
        && left.custom_shows.shows == right.custom_shows.shows
        && left.sections == right.sections
}

/// # Errors
///
/// Returns an error if the operation fails.
pub fn find_custom_show(package: &OpcPackage, id: u32) -> Result<Option<Show>> {
    Ok(
        load(package, Limits::from_source_limits(package.read_limits()))?
            .custom_shows
            .get_by_id(id)
            .cloned(),
    )
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add_custom_show(package: &mut OpcPackage, show: Show) -> Result<()> {
    mutate(package, |graph| {
        if graph.custom_shows.get_by_id(show.id).is_some()
            || graph.custom_shows.get_by_name(&show.name).is_some()
        {
            return Err(invalid("duplicate custom-show ID or name"));
        }
        graph.custom_shows.add(show);
        Ok(())
    })
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn update_custom_show(
    package: &mut OpcPackage,
    id: u32,
    replacement: Show,
) -> Result<()> {
    mutate(package, |graph| {
        graph.custom_shows.replace_by_id(id, replacement)
    })
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn remove_custom_show(package: &mut OpcPackage, id: u32) -> Result<bool> {
    let mut graph = load(package, Limits::from_source_limits(package.read_limits()))?;
    let removed = graph.custom_shows.remove_by_id(id).is_some();
    if removed {
        store(package, &graph)?;
    }
    Ok(removed)
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn reorder_custom_shows(package: &mut OpcPackage, ordered_ids: &[u32]) -> Result<()> {
    mutate(package, |graph| graph.custom_shows.reorder(ordered_ids))
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add_custom_show_slide(
    package: &mut OpcPackage,
    show_id: u32,
    slide_id: u32,
) -> Result<()> {
    mutate(package, |graph| {
        require_slide(graph, slide_id)?;
        let show = graph
            .custom_shows
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        if show.slide_ids.contains(&slide_id) {
            return Err(invalid("duplicate custom-show slide membership"));
        }
        show.slide_ids.push(slide_id);
        Ok(())
    })
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn remove_custom_show_slide(
    package: &mut OpcPackage,
    show_id: u32,
    slide_id: u32,
) -> Result<bool> {
    let mut graph = load(package, Limits::from_source_limits(package.read_limits()))?;
    let show = graph
        .custom_shows
        .get_by_id_mut(show_id)
        .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
    let Some(offset) = show.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    show.slide_ids.remove(offset);
    store(package, &graph)?;
    Ok(true)
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn reorder_custom_show_slides(
    package: &mut OpcPackage,
    show_id: u32,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    mutate(package, |graph| {
        let show = graph
            .custom_shows
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        require_permutation(&show.slide_ids, ordered_slide_ids, "custom-show slide")?;
        show.slide_ids = ordered_slide_ids.to_vec();
        Ok(())
    })
}

/// # Errors
///
/// Returns an error if the operation fails.
pub fn find_section(package: &OpcPackage, id: &str) -> Result<Option<Section>> {
    Ok(
        load(package, Limits::from_source_limits(package.read_limits()))?
            .sections
            .get_by_id(id)
            .cloned(),
    )
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add_section(package: &mut OpcPackage, mut section: Section) -> Result<String> {
    if section.id.is_none() {
        section.id = Some(new_guid());
    }
    let id = section
        .id
        .clone()
        .ok_or_else(|| invalid("section ID assignment failed"))?;
    let retained = id.clone();
    mutate(package, move |graph| {
        if graph.sections.get_by_id(&id).is_some() {
            return Err(invalid(format!("section {id} already exists")));
        }
        graph.sections.add_section(section);
        graph.sections.sort_slide_membership(
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
        );
        Ok(())
    })?;
    Ok(retained)
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn update_section(
    package: &mut OpcPackage,
    id: &str,
    replacement: Section,
) -> Result<()> {
    let id = id.to_owned();
    let limits = Limits::from_source_limits(package.read_limits());
    mutate(package, move |graph| {
        graph.sections.replace_by_id(&id, replacement, limits)?;
        graph.sections.sort_slide_membership(
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
        );
        Ok(())
    })
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn remove_section(package: &mut OpcPackage, id: &str) -> Result<bool> {
    let mut graph = load(package, Limits::from_source_limits(package.read_limits()))?;
    let removed = graph.sections.remove_by_id(id).is_some();
    if removed {
        store(package, &graph)?;
    }
    Ok(removed)
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn reorder_sections(package: &mut OpcPackage, ordered_ids: &[String]) -> Result<()> {
    mutate(package, |graph| graph.sections.reorder(ordered_ids))
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn add_section_slide(
    package: &mut OpcPackage,
    section_id: &str,
    slide_id: u32,
) -> Result<()> {
    let section_id = section_id.to_owned();
    mutate(package, move |graph| {
        require_slide(graph, slide_id)?;
        if graph
            .sections
            .sections()
            .iter()
            .any(|section| section.slide_ids.contains(&slide_id))
        {
            return Err(invalid("slide already belongs to a section"));
        }
        graph
            .sections
            .get_by_id_mut(&section_id)
            .ok_or_else(|| invalid(format!("section {section_id} was not found")))?
            .slide_ids
            .push(slide_id);
        graph.sections.sort_slide_membership(
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
        );
        Ok(())
    })
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn remove_section_slide(
    package: &mut OpcPackage,
    section_id: &str,
    slide_id: u32,
) -> Result<bool> {
    let mut graph = load(package, Limits::from_source_limits(package.read_limits()))?;
    let section = graph
        .sections
        .get_by_id_mut(section_id)
        .ok_or_else(|| invalid(format!("section {section_id} was not found")))?;
    let Some(offset) = section.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    section.slide_ids.remove(offset);
    store(package, &graph)?;
    Ok(true)
}

/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn reorder_section_slides(
    package: &mut OpcPackage,
    section_id: &str,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    let section_id = section_id.to_owned();
    mutate(package, move |graph| {
        let section = graph
            .sections
            .get_by_id_mut(&section_id)
            .ok_or_else(|| invalid(format!("section {section_id} was not found")))?;
        require_permutation(&section.slide_ids, ordered_slide_ids, "section slide")?;
        section.slide_ids = ordered_slide_ids.to_vec();
        Ok(())
    })
}

fn mutate<F>(package: &mut OpcPackage, operation: F) -> Result<()>
where
    F: FnOnce(&mut Graph) -> Result<()>,
{
    let mut graph = load(package, Limits::from_source_limits(package.read_limits()))?;
    operation(&mut graph)?;
    store(package, &graph)
}

fn parse_structure_blob(
    package: &OpcPackage,
    xml: &[u8],
    strict_references: bool,
    limits: Limits,
) -> Result<Graph> {
    if xml.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            xml.len(),
            limits.max_input_bytes(),
        ));
    }
    let processed = process_ooxml(xml, limits)?;
    if processed.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            processed.len(),
            limits.max_output_bytes(),
        ));
    }
    let (raw_slides, raw_shows) = parse_core(processed.as_ref(), limits)?;
    let presentation = package.main_document_part()?;
    let mut slides = Vec::with_capacity(raw_slides.len());
    for (slide_id, relationship_id) in raw_slides {
        let relationship = presentation.rels().get(&relationship_id).ok_or_else(|| {
            invalid(format!(
                "orphan presentation slide relationship {relationship_id}"
            ))
        })?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), SLIDE_REL | SLIDE_REL_STRICT)
        {
            return Err(invalid(format!(
                "relationship {relationship_id} is not an internal slide relationship"
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type()
            != "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
        {
            return Err(invalid(format!(
                "relationship {relationship_id} targets a non-slide part"
            )));
        }
        slides.push(Reference {
            slide_id,
            relationship_id,
            part_name: target.to_string(),
        });
    }
    let custom_shows = resolve_custom_shows(raw_shows, &slides, strict_references)?;
    let graph = Graph {
        slides,
        custom_shows,
        sections: SectionList::from_xml(xml, limits)?,
    };
    if strict_references {
        validate_graph(package, &graph, limits)?;
    }
    Ok(graph)
}

fn resolve_custom_shows(
    raw_shows: Vec<RawShow>,
    slides: &[Reference],
    strict_references: bool,
) -> Result<ShowList> {
    let rel_to_id = slides
        .iter()
        .map(|slide| (slide.relationship_id.as_str(), slide.slide_id))
        .collect::<HashMap<_, _>>();
    let mut custom_shows = ShowList::new();
    for raw in raw_shows {
        let mut show = Show::new(raw.id, raw.name);
        for relationship_id in raw.relationship_ids {
            if let Some(slide_id) = rel_to_id.get(relationship_id.as_str()) {
                show.slide_ids.push(*slide_id);
            } else if strict_references {
                return Err(invalid(format!(
                    "custom show {} has orphan slide relationship {relationship_id}",
                    show.id
                )));
            }
        }
        custom_shows.add(show);
    }
    Ok(custom_shows)
}

#[derive(Default)]
pub(super) struct RawShow {
    pub(super) id: u32,
    pub(super) name: String,
    pub(super) relationship_ids: Vec<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum CoreElement {
    Presentation,
    SlideIdList,
    SlideId,
    CustomShowList,
    CustomShow,
    SlideList,
    Slide,
    Other,
}

#[derive(Debug)]
struct Attribute {
    namespace: Option<Vec<u8>>,
    local: Vec<u8>,
    value: String,
}

#[allow(
    clippy::type_complexity,
    reason = "raw-show tuple mirrors the persisted XML structure"
)]
pub(super) fn parse_core(xml: &[u8], limits: Limits) -> Result<(Vec<(u32, String)>, Vec<RawShow>)> {
    if xml.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            xml.len(),
            limits.max_output_bytes(),
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    // Keep raw text events intact until their selected string budget has been
    // admitted.  `parse_core` ignores text semantically, so disabling
    // quick-xml trimming preserves the graph result while preventing a long
    // whitespace-only event from being shortened before preflight.
    reader.config_mut().trim_text(false);
    let mut ancestors = Vec::<CoreElement>::new();
    let mut slides = Vec::new();
    let mut shows = Vec::new();
    let mut current_show: Option<RawShow> = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        let borrowed = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        // Admission must happen while quick-xml still borrows the source.
        // Calling `into_owned` first would copy attacker-controlled names,
        // attributes, text, or references before the selected string policy
        // had a chance to reject them.
        preflight_core_event(&borrowed, limits)?;
        let event = borrowed.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                resource(nodes, ancestors.len(), limits)?;
                let kind = core_element(&namespace, element.name())?;
                if is_direct_slide_reference(&ancestors, kind) {
                    if slides.len() >= limits.max_slide_references() {
                        return Err(limit(
                            crate::opened::CodecLimitKind::SlideReferences,
                            slides.len() + 1,
                            limits.max_slide_references(),
                        ));
                    }
                    slides.push(parse_slide_reference(&element, decoder, &reader, limits)?);
                } else if kind == CoreElement::CustomShow {
                    if ancestors.as_slice()
                        != [CoreElement::Presentation, CoreElement::CustomShowList]
                    {
                        return Err(invalid("custom show has invalid presentation ancestry"));
                    }
                    if current_show.is_some() {
                        return Err(invalid("nested custom shows are rejected"));
                    }
                    current_show = Some(parse_show(&element, decoder, &reader, limits)?);
                }
                ancestors.push(kind);
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                resource(nodes, ancestors.len(), limits)?;
                let kind = core_element(&namespace, element.name())?;
                if is_direct_slide_reference(&ancestors, kind) {
                    if slides.len() >= limits.max_slide_references() {
                        return Err(limit(
                            crate::opened::CodecLimitKind::SlideReferences,
                            slides.len() + 1,
                            limits.max_slide_references(),
                        ));
                    }
                    slides.push(parse_slide_reference(&element, decoder, &reader, limits)?);
                } else if kind == CoreElement::Slide
                    && ancestors.as_slice()
                        == [
                            CoreElement::Presentation,
                            CoreElement::CustomShowList,
                            CoreElement::CustomShow,
                            CoreElement::SlideList,
                        ]
                {
                    let relationship_id = relationship_id(&element, decoder, &reader, limits)?;
                    current_show
                        .as_mut()
                        .ok_or_else(|| invalid("custom-show slide appears outside a custom show"))?
                        .relationship_ids
                        .push(relationship_id);
                    if current_show.as_ref().is_some_and(|show| {
                        show.relationship_ids.len() > limits.max_slide_references()
                    }) {
                        return Err(limit(
                            crate::opened::CodecLimitKind::SlideReferences,
                            current_show
                                .as_ref()
                                .map_or(usize::MAX, |show| show.relationship_ids.len()),
                            limits.max_slide_references(),
                        ));
                    }
                }
            },
            Event::End(element) => {
                let kind = core_element(&namespace, element.name())?;
                let open = ancestors
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                if open != kind {
                    return Err(invalid("mismatched presentation element"));
                }
                if kind == CoreElement::CustomShow {
                    if shows.len() >= limits.max_custom_shows() {
                        return Err(limit(
                            crate::opened::CodecLimitKind::CustomShows,
                            shows.len() + 1,
                            limits.max_custom_shows(),
                        ));
                    }
                    shows.push(
                        current_show
                            .take()
                            .ok_or_else(|| invalid("missing custom show"))?,
                    );
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !ancestors.is_empty() || current_show.is_some() {
        return Err(invalid("unterminated presentation structure"));
    }
    Ok((slides, shows))
}

fn preflight_core_event(event: &Event<'_>, limits: Limits) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            check_raw_name(element.name().as_ref(), limits)?;
            let mut count = 0usize;
            for attribute in element.checked_attributes() {
                let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                count = count.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlAttributes,
                        usize::MAX,
                        limits.max_attributes(),
                    )
                })?;
                if count > limits.max_attributes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlAttributes,
                        count,
                        limits.max_attributes(),
                    ));
                }
                if attribute.key.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlAttributeBytes,
                        attribute.key.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                if attribute.value.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlAttributeBytes,
                        attribute.value.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
            }
        },
        Event::End(element) => check_raw_name(element.name().as_ref(), limits)?,
        Event::Text(text) => check_raw_text(
            text.as_ref(),
            limits,
            crate::opened::CodecLimitKind::TextBytes,
        )?,
        Event::CData(text) => check_raw_text(
            text.as_ref(),
            limits,
            crate::opened::CodecLimitKind::TextBytes,
        )?,
        Event::Comment(text) => check_raw_text(
            text.as_ref(),
            limits,
            crate::opened::CodecLimitKind::TextBytes,
        )?,
        Event::GeneralRef(reference) => check_raw_text(
            reference.as_ref(),
            limits,
            crate::opened::CodecLimitKind::ReferenceBytes,
        )?,
        Event::Decl(declaration) => check_raw_text(
            declaration.as_ref(),
            limits,
            crate::opened::CodecLimitKind::XmlBytes,
        )?,
        Event::DocType(_) | Event::PI(_) => {
            // Reject these borrowed events before `into_owned`; their payload
            // is not allowed to reach an owned allocation on the error path.
            return Err(invalid("DTDs and processing instructions are rejected"));
        },
        Event::Eof => {},
    }
    Ok(())
}

fn check_raw_name(name: &[u8], limits: Limits) -> Result<()> {
    if name.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::NameBytes,
            name.len(),
            limits.max_string_bytes(),
        ));
    }
    Ok(())
}

fn check_raw_text(text: &[u8], limits: Limits, kind: crate::opened::CodecLimitKind) -> Result<()> {
    if text.len() > limits.max_string_bytes() {
        return Err(limit(kind, text.len(), limits.max_string_bytes()));
    }
    Ok(())
}

fn core_element(namespace: &ResolveResult<'_>, name: QName<'_>) -> Result<CoreElement> {
    let presentationml = match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            matches!(*value, value if value == P.as_bytes() || value == PS.as_bytes())
        },
        ResolveResult::Unbound => false,
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unresolved presentation XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    if !presentationml {
        return Ok(CoreElement::Other);
    }
    Ok(match name.local_name().as_ref() {
        b"presentation" => CoreElement::Presentation,
        b"sldIdLst" => CoreElement::SlideIdList,
        b"sldId" => CoreElement::SlideId,
        b"custShowLst" => CoreElement::CustomShowList,
        b"custShow" => CoreElement::CustomShow,
        b"sldLst" => CoreElement::SlideList,
        b"sld" => CoreElement::Slide,
        _ => CoreElement::Other,
    })
}

fn is_direct_slide_reference(ancestors: &[CoreElement], kind: CoreElement) -> bool {
    kind == CoreElement::SlideId
        && ancestors == [CoreElement::Presentation, CoreElement::SlideIdList]
}

fn parse_show(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<RawShow> {
    let attributes = attributes_ns(element, decoder, reader, limits)?;
    let id = required(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_err| invalid("invalid custom-show ID"))?;
    let name = required(&attributes, "name")?;
    if name.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::StringBytes,
            name.len(),
            limits.max_string_bytes(),
        ));
    }
    if name.is_empty() {
        return Err(invalid("custom-show name cannot be empty"));
    }
    Ok(RawShow {
        id,
        name: name.to_owned(),
        relationship_ids: Vec::new(),
    })
}

fn parse_slide_reference(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<(u32, String)> {
    let attributes = attributes_ns(element, decoder, reader, limits)?;
    let id = required_unqualified(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_err| invalid("invalid presentation slide ID"))?;
    if !(256..=MAX_SLIDE_ID).contains(&id) {
        return Err(invalid("presentation slide ID is outside 256..=2147483647"));
    }
    let relationship_id = required_qualified(&attributes, "id")?;
    if relationship_id.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::RelationshipBytes,
            relationship_id.len(),
            limits.max_string_bytes(),
        ));
    }
    Ok((id, relationship_id.to_owned()))
}

fn relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<String> {
    let attributes = attributes_ns(element, decoder, reader, limits)?;
    let relationship_id = required_qualified(&attributes, "id")?;
    if relationship_id.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::RelationshipBytes,
            relationship_id.len(),
            limits.max_string_bytes(),
        ));
    }
    Ok(relationship_id.to_owned())
}

fn attributes_ns(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<Vec<Attribute>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if values.len() >= limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                values.len() + 1,
                limits.max_attributes(),
            ));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => {
                if value.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NamespaceBytes,
                        value.len(),
                        limits.max_string_bytes(),
                    ));
                }
                Some(value.to_vec())
            },
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "unresolved presentation XML attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if local.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::NameBytes,
                local.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let local = local.as_ref().to_vec();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(invalid("duplicate presentation XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
        values.push(Attribute {
            namespace,
            local,
            value: value.into_owned(),
        });
    }
    Ok(values)
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if values.len() >= limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                values.len() + 1,
                limits.max_attributes(),
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if name.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::NameBytes,
                name.len(),
                limits.max_string_bytes(),
            ));
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let name = name.to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate presentation XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
        values.push((name, value.into_owned()));
    }
    Ok(values)
}

fn required<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local == local.as_bytes())
        .map(|attribute| attribute.value.as_str())
        .ok_or_else(|| invalid(format!("missing attribute '{local}'")))
}

fn required_unqualified<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    required(values, local)
}

fn required_qualified<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|attribute| {
            attribute.local == local.as_bytes()
                && attribute.namespace.as_deref().is_some_and(|namespace| {
                    namespace == R.as_bytes() || namespace == RS.as_bytes()
                })
        })
        .map(|attribute| attribute.value.as_str())
        .ok_or_else(|| invalid(format!("missing relationship attribute '{local}'")))
}

fn validate_graph(package: &OpcPackage, graph: &Graph, limits: Limits) -> Result<()> {
    validate_graph_shape(graph, limits)?;

    let presentation = package.main_document_part()?;
    for slide in &graph.slides {
        let relationship = presentation
            .rels()
            .get(&slide.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "orphan slide relationship {}",
                    slide.relationship_id
                ))
            })?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), SLIDE_REL | SLIDE_REL_STRICT)
            || relationship.target_partname()?.as_str() != slide.part_name
        {
            return Err(invalid("presentation slide relationship mismatch"));
        }
    }
    Ok(())
}

fn validate_graph_shape(graph: &Graph, limits: Limits) -> Result<()> {
    let mut slide_ids = HashSet::new();
    let mut rel_ids = HashSet::new();
    let mut part_names = HashSet::new();
    for slide in &graph.slides {
        if !(256..=MAX_SLIDE_ID).contains(&slide.slide_id) || !slide_ids.insert(slide.slide_id) {
            return Err(invalid("invalid or duplicate presentation slide ID"));
        }
        if !rel_ids.insert(slide.relationship_id.as_str())
            || !part_names.insert(slide.part_name.as_str())
        {
            return Err(invalid(
                "duplicate presentation slide relationship or target",
            ));
        }
    }
    let mut show_ids = HashSet::new();
    let mut show_names = HashSet::new();
    if graph.custom_shows.shows.len() > limits.max_custom_shows() {
        return Err(limit(
            crate::opened::CodecLimitKind::CustomShows,
            graph.custom_shows.shows.len(),
            limits.max_custom_shows(),
        ));
    }
    let mut total_show_slides = 0usize;
    for show in &graph.custom_shows.shows {
        if show.name.is_empty()
            || !show_ids.insert(show.id)
            || !show_names.insert(show.name.as_str())
        {
            return Err(invalid("duplicate or empty custom-show identity"));
        }
        let mut membership = HashSet::new();
        for slide_id in &show.slide_ids {
            if !slide_ids.contains(slide_id) || !membership.insert(*slide_id) {
                return Err(invalid("duplicate or orphan custom-show slide reference"));
            }
        }
        total_show_slides = total_show_slides
            .checked_add(show.slide_ids.len())
            .ok_or_else(|| invalid("custom-show slide reference count overflow"))?;
        if total_show_slides > limits.max_slide_references() {
            return Err(limit(
                crate::opened::CodecLimitKind::SlideReferences,
                total_show_slides,
                limits.max_slide_references(),
            ));
        }
    }
    let positions = graph
        .slides
        .iter()
        .enumerate()
        .map(|(offset, slide)| (slide.slide_id, offset))
        .collect::<HashMap<_, _>>();
    let mut section_ids = HashSet::new();
    let mut section_membership = HashSet::new();
    for section in graph.sections.sections() {
        let id = section
            .id
            .as_deref()
            .ok_or_else(|| invalid("mutable presentation section requires a stable GUID"))?;
        if !section_ids.insert(id) {
            return Err(invalid("duplicate section GUID"));
        }
        let mut previous = None;
        for slide_id in &section.slide_ids {
            let position = positions
                .get(slide_id)
                .copied()
                .ok_or_else(|| invalid("orphan section slide reference"))?;
            if !section_membership.insert(*slide_id) {
                return Err(invalid("slide belongs to more than one section"));
            }
            if previous.is_some_and(|value| value >= position) {
                return Err(invalid("section membership is not in presentation order"));
            }
            previous = Some(position);
        }
    }
    if graph.sections.sections().len() > limits.max_sections() {
        return Err(limit(
            crate::opened::CodecLimitKind::Sections,
            graph.sections.sections().len(),
            limits.max_sections(),
        ));
    }
    graph.sections.to_xml(limits)?;
    Ok(())
}

fn write_section_extension(
    sections: &SectionList,
    p_namespace: &str,
    limits: Limits,
) -> Result<String> {
    let full = sections.to_xml(limits)?;
    if full.is_empty() {
        return Ok(String::new());
    }
    let inner = full
        .strip_prefix("<p:extLst>")
        .and_then(|value| value.strip_suffix("</p:extLst>"))
        .ok_or_else(|| invalid("invalid serialized section extension"))?;
    let Some(prefix) = inner.strip_prefix("<p:ext ") else {
        return Err(invalid("invalid serialized section extension"));
    };
    let mut output = BoundedString::new(limits, "serialized presentation sections");
    output.push_str("<p:ext xmlns:p=\"")?;
    output.push_escaped_xml(p_namespace)?;
    output.push_str("\" ")?;
    output.push_str(prefix)?;
    Ok(output.finish())
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
    close_start: usize,
    local: String,
}

struct Frame {
    start: usize,
    local: String,
    direct: bool,
    #[cfg(test)]
    target_custom: bool,
    target_section: bool,
}

struct XmlLayout {
    root_close: usize,
    direct: Vec<Span>,
    #[cfg(test)]
    custom: Vec<Span>,
    section_extensions: Vec<Span>,
}

fn scan_layout(xml: &[u8], limits: Limits) -> Result<XmlLayout> {
    if xml.len() > limits.max_input_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::InputBytes,
            xml.len(),
            limits.max_input_bytes(),
        ));
    }
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut stack = Vec::<Frame>::new();
    let mut direct = Vec::new();
    #[cfg(test)]
    let mut custom = Vec::new();
    let mut section_extensions = Vec::new();
    let mut root_close = None;
    let mut nodes = 0usize;
    loop {
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML source position exceeds usize"))?;
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                resource(nodes, stack.len(), limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.name().as_ref())?;
                let target_section = local == "ext"
                    && attributes(&element, decoder, limits)?
                        .iter()
                        .any(|(name, value)| name == "uri" && value == SECTION_URI);
                stack.push(Frame {
                    start: before,
                    #[cfg(test)]
                    target_custom: local == "custShowLst",
                    target_section,
                    direct: stack.len() == 1,
                    local,
                });
            },
            Ok(Event::Empty(element)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                resource(nodes, stack.len(), limits)?;
                if element.name().as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::NameBytes,
                        element.name().as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let local = local_name(element.name().as_ref())?;
                let span = Span {
                    start: before,
                    end: origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?,
                    close_start: before,
                    local: local.clone(),
                };
                if stack.len() == 1 {
                    direct.push(span.clone());
                }
                #[cfg(test)]
                if local == "custShowLst" {
                    custom.push(span.clone());
                }
                if local == "ext"
                    && attributes(&element, decoder, limits)?
                        .iter()
                        .any(|(name, value)| name == "uri" && value == SECTION_URI)
                {
                    section_extensions.push(span);
                }
            },
            Ok(Event::End(_)) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                let span = Span {
                    start: frame.start,
                    end: origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?,
                    close_start: before,
                    local: frame.local,
                };
                if frame.direct {
                    direct.push(span.clone());
                }
                #[cfg(test)]
                if frame.target_custom {
                    custom.push(span.clone());
                }
                if frame.target_section {
                    section_extensions.push(span.clone());
                }
                if stack.is_empty() {
                    root_close = Some(before);
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    direct.sort_by_key(|span| span.start);
    Ok(XmlLayout {
        root_close: root_close.ok_or_else(|| invalid("missing presentation root close"))?,
        direct,
        #[cfg(test)]
        custom,
        section_extensions,
    })
}

#[cfg(test)]
fn patch_custom_shows(xml: &[u8], replacement: &[u8], limits: Limits) -> Result<Vec<u8>> {
    let layout = scan_layout(xml, limits)?;
    if !layout.custom.is_empty() {
        return replace_spans(xml, &layout.custom, replacement, limits);
    }
    if replacement.is_empty() {
        if xml.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                xml.len(),
                limits.max_output_bytes(),
            ));
        }
        return Ok(xml.to_vec());
    }
    let insert = layout
        .direct
        .iter()
        .find(|span| schema_rank(&span.local).is_some_and(|rank| rank > 8))
        .map_or(layout.root_close, |span| span.start);
    insert_bytes(xml, insert, replacement, limits)
}

fn patch_sections(
    xml: &[u8],
    replacement: &[u8],
    p_namespace: &str,
    limits: Limits,
) -> Result<Vec<u8>> {
    let layout = scan_layout(xml, limits)?;
    if !layout.section_extensions.is_empty() {
        return replace_spans(xml, &layout.section_extensions, replacement, limits);
    }
    if replacement.is_empty() {
        if xml.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                xml.len(),
                limits.max_output_bytes(),
            ));
        }
        return Ok(xml.to_vec());
    }
    if let Some(ext_list) = layout.direct.iter().find(|span| span.local == "extLst") {
        if ext_list.close_start == ext_list.start {
            let wrapped = wrap_extension_list(replacement, p_namespace, limits)?;
            return replace_spans(
                xml,
                std::slice::from_ref(ext_list),
                wrapped.as_bytes(),
                limits,
            );
        }
        return insert_bytes(xml, ext_list.close_start, replacement, limits);
    }
    let wrapped = wrap_extension_list(replacement, p_namespace, limits)?;
    insert_bytes(xml, layout.root_close, wrapped.as_bytes(), limits)
}

fn wrap_extension_list(replacement: &[u8], p_namespace: &str, limits: Limits) -> Result<String> {
    let replacement =
        std::str::from_utf8(replacement).map_err(|error| Error::Xml(error.to_string()))?;
    let mut wrapped = BoundedString::new(limits, "patched presentation structure");
    wrapped.push_str("<p:extLst xmlns:p=\"")?;
    wrapped.push_escaped_xml(p_namespace)?;
    wrapped.push_str("\">")?;
    wrapped.push_str(replacement)?;
    wrapped.push_str("</p:extLst>")?;
    Ok(wrapped.finish())
}

fn replace_spans(
    xml: &[u8],
    spans: &[Span],
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    let mut ordered = spans.to_vec();
    ordered.sort_by_key(|span| span.start);
    for pair in ordered.windows(2) {
        if pair[0].end > pair[1].start {
            return Err(invalid("overlapping presentation XML patch ranges"));
        }
    }
    let mut removed = 0usize;
    for span in &ordered {
        if span.start > span.end || span.end > xml.len() {
            return Err(invalid("presentation XML patch range is outside source"));
        }
        removed = removed
            .checked_add(span.end - span.start)
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation structure",
            })?;
    }
    let output_len = checked_output_len(
        xml.len(),
        removed,
        replacement
            .len()
            .checked_mul(ordered.len())
            .ok_or(Error::CapacityOverflow {
                resource: "patched presentation structure",
            })?,
        limits,
        "patched presentation structure",
    )?;
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched presentation structure",
            source,
        })?;
    let mut cursor = 0usize;
    for span in ordered {
        output.extend_from_slice(&xml[cursor..span.start]);
        output.extend_from_slice(replacement);
        cursor = span.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn insert_bytes(xml: &[u8], offset: usize, value: &[u8], limits: Limits) -> Result<Vec<u8>> {
    if offset > xml.len() {
        return Err(invalid(
            "presentation XML insertion offset is outside source",
        ));
    }
    let output_len = xml
        .len()
        .checked_add(value.len())
        .ok_or(Error::CapacityOverflow {
            resource: "patched presentation structure",
        })?;
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            output_len,
            limits.max_output_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched presentation structure",
            source,
        })?;
    output.extend_from_slice(&xml[..offset]);
    output.extend_from_slice(value);
    output.extend_from_slice(&xml[offset..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

#[cfg(test)]
fn schema_rank(local: &str) -> Option<usize> {
    [
        "sldMasterIdLst",
        "notesMasterIdLst",
        "handoutMasterIdLst",
        "sldIdLst",
        "sldSz",
        "notesSz",
        "smartTags",
        "embeddedFontLst",
        "custShowLst",
        "photoAlbum",
        "custDataLst",
        "kinsoku",
        "defaultTextStyle",
        "modifyVerifier",
        "extLst",
    ]
    .iter()
    .position(|name| *name == local)
}

fn document_namespaces(xml: &[u8]) -> (&'static str, &'static str) {
    if xml.windows(PS.len()).any(|window| window == PS.as_bytes()) {
        (PS, RS)
    } else {
        (P, R)
    }
}

fn require_slide(graph: &Graph, slide_id: u32) -> Result<()> {
    if graph.slides.iter().any(|slide| slide.slide_id == slide_id) {
        Ok(())
    } else {
        Err(invalid(format!("slide {slide_id} was not found")))
    }
}

fn require_permutation(expected: &[u32], actual: &[u32], what: &str) -> Result<()> {
    let expected_set = expected.iter().copied().collect::<HashSet<_>>();
    let actual_set = actual.iter().copied().collect::<HashSet<_>>();
    if expected_set != actual_set || expected.len() != actual.len() {
        Err(invalid(format!("{what} reorder is not a permutation")))
    } else {
        Ok(())
    }
}

fn require_presentation(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("main document is not a PowerPoint presentation"))
    }
}

fn resource(nodes: usize, depth: usize, limits: Limits) -> Result<()> {
    if nodes > limits.max_nodes() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlNodes,
            nodes,
            limits.max_nodes(),
        ));
    }
    if depth >= limits.max_depth() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlDepth,
            depth + 1,
            limits.max_depth(),
        ));
    }
    Ok(())
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(|error| Error::Xml(error.to_string()))?;
    Ok(name
        .rsplit_once(':')
        .map_or(name, |(_, local)| local)
        .to_owned())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::opened::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::opened::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn malformed_xml_is_rejected_and_mce_branches_patch_without_overlap() {
        assert!(parse_core(b"<!DOCTYPE x><p:presentation/>", Limits::DEFAULT).is_err());
        assert!(parse_core(b"<p:presentation><p:sldIdLst>", Limits::DEFAULT).is_err());

        let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="p"><p:custShowLst><p:custShow name="old" id="1"><p:sldLst/></p:custShow></p:custShowLst></mc:Choice><mc:Fallback><p:custShowLst/></mc:Fallback></mc:AlternateContent><!--keep--></p:presentation>"#;
        let replacement = br#"<p:custShowLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let patched = patch_custom_shows(xml, replacement, Limits::DEFAULT).unwrap();
        assert_eq!(
            patched
                .windows(replacement.len())
                .filter(|window| *window == replacement)
                .count(),
            2
        );
        assert!(
            patched
                .windows(b"<!--keep-->".len())
                .any(|window| window == b"<!--keep-->")
        );
    }

    #[test]
    fn non_empty_slide_ids_are_recorded_once_and_indirect_lookalikes_are_ignored() {
        let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:future"><p:sldIdLst><p:sldId id="256" r:id="rIdOne"><p:extLst><p:ext uri="{opaque}"><x:future><p:sldId id="999" r:id="rIdNested"/></x:future></p:ext></p:extLst></p:sldId><p:extLst><p:sldId id="998" r:id="rIdIndirect"/></p:extLst><p:sldId id="257" r:id="rIdTwo"/></p:sldIdLst><p:extLst><p:sldId id="997" r:id="rIdOutside"/></p:extLst></p:presentation>"#;

        let (slides, shows) = parse_core(xml, Limits::DEFAULT).unwrap();

        assert_eq!(
            slides,
            vec![(256, "rIdOne".to_owned()), (257, "rIdTwo".to_owned())]
        );
        assert!(shows.is_empty());
    }
}
