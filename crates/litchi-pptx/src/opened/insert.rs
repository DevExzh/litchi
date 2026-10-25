//! Source-preserving insertion through native authoring and OPC metadata plans.
use super::model::invalid;
use super::{Patch, Snapshot};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, SourceLimits, TargetMode};
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

fn refuse(reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "insert_slide",
        reason,
    )
}
fn source_refusal(error: Error) -> Error {
    match error {
        // A typed refusal is an unsupported source route.  Preserve invalid
        // XML/OPC validation failures as malformed input so they retain the
        // corrupt-file classification at the facade.
        Error::Refused(_) => refuse("the source is outside the supported slide topology"),
        other => other,
    }
}
fn strict(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (ns, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return match ns {
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(false)
                    },
                    ResolveResult::Bound(Namespace(value))
                        if value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE =>
                    {
                        Ok(true)
                    },
                    _ => Err(refuse("unsupported presentation namespace")),
                };
            },
            Event::Eof => return Err(invalid("missing presentation root")),
            _ => {},
        }
    }
}

pub(super) fn plan(source: &Snapshot, position: usize, layout_name: &str) -> Result<Patch> {
    if position > source.slides.len() {
        return Err(Error::SlideIndexOutOfBounds {
            index: position,
            len: source.slides.len().checked_add(1).unwrap_or(usize::MAX),
        });
    }
    if source.slides.len() >= crate::parts::MAX_SLIDES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::Slides),
            observed: source.slides.len().checked_add(1).unwrap_or(usize::MAX),
            maximum: crate::parts::MAX_SLIDES,
        });
    }
    if source.package.part_count() >= source.limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Parts,
            observed: source
                .package
                .part_count()
                .checked_add(1)
                .unwrap_or(usize::MAX),
            maximum: source.limits.max_parts(),
        });
    }
    if layout_name.len() > source.limits.max_text_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::NameBytes),
            observed: layout_name.len(),
            maximum: source.limits.max_text_bytes(),
        });
    }
    if super::copy_plan::has_signature_infrastructure(&source.package) {
        return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
    }
    super::copy_plan::reject_unknown_non_part_members(&source.package, "slide insertion source")
        .map_err(source_refusal)?;
    let main = source.package.get_part(&source.presentation_name)?;
    // Both old and new presentation payloads must fit the durable patch. This
    // lower-bound admission precedes XML scans and lexical output allocation.
    let main_patch_charge = main
        .blob()
        .len()
        .checked_mul(2)
        .ok_or_else(|| Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: usize::MAX,
            maximum: source.limits.max_patch_bytes(),
        })?;
    if main_patch_charge > source.limits.max_patch_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: main_patch_charge,
            maximum: source.limits.max_patch_bytes(),
        });
    }
    super::copy_plan::reject_mce(main.blob(), "presentation owner").map_err(source_refusal)?;
    let main_text =
        std::str::from_utf8(main.blob()).map_err(|_| invalid("presentation XML is not UTF-8"))?;
    if crate::presentation_properties::metadata::protection::Settings::parse_xml(
        main_text,
        crate::presentation_properties::metadata::protection::Limits::from_source_limits(
            source.package.read_limits(),
        ),
    )?
    .is_protected()
    {
        return Err(refuse(
            "the presentation has an active modify-password verifier",
        ));
    }
    if source
        .slides
        .iter()
        .any(|slide| !(256..=2_147_483_647).contains(&slide.id))
    {
        return Err(refuse(
            "an existing slide ID is outside the PresentationML range",
        ));
    }
    crate::master_layout::validate_master_layout_graph(&source.package)?;
    let presentation = crate::Presentation::new_with_semantic_limits(
        crate::parts::PresentationPart::from_part(main)?,
        &source.package,
        source.owner_state.semantic_limits.clone(),
    );
    let mut layout = None;
    let mut matches = 0;
    for candidate in presentation.slide_layouts()? {
        if candidate.name()? == layout_name {
            matches += 1;
            layout = Some(candidate.part().part().partname().clone());
        }
    }
    if matches > 1 {
        return Err(Error::AmbiguousSlideLayoutName {
            name: layout_name.to_owned(),
            matches,
        });
    }
    let layout = layout.ok_or_else(|| Error::SlideLayoutNameNotFound(layout_name.to_owned()))?;
    let strict = strict(main.blob())?;
    let id = super::copy_plan::next_slide_id(&source.slides).map_err(source_refusal)?;
    let relationship_id =
        super::copy_plan::next_relationship_id(main.rels()).map_err(source_refusal)?;
    super::remove_plan::validate_insertion_owner(main.blob(), &relationship_id)
        .map_err(source_refusal)?;
    let name = available_name(source)?;
    let maximum = source.limits.max_patch_bytes();
    let policy = SourceLimits::builder()
        .max_part_bytes(maximum as u64)?
        .max_entry_bytes(maximum as u64)?
        .max_content_types_bytes(maximum)?
        .max_relationship_xml_bytes(maximum)?
        .max_content_type_mappings(source.limits.max_parts().saturating_add(2))?
        .max_relationships_per_part(source.limits.max_parts())?
        .build()?;
    let types = source.package.source_content_types(policy)?;
    let rels = source
        .package
        .source_relationships(&source.presentation_name, policy)?;
    let types_plan = types.plan_edit(
        &[litchi_opc::ContentTypeEdit {
            part_name: &name,
            content_type: ct::PML_SLIDE,
        }],
        &[],
        policy,
    )?;
    let slide_type = if strict {
        "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
    } else {
        rt::SLIDE
    };
    let layout_type = if strict {
        "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout"
    } else {
        rt::SLIDE_LAYOUT
    };
    let target = name.relative_ref(source.presentation_name.base_uri());
    let rels_plan = rels.plan_edit(
        &[litchi_opc::RelationshipEdit {
            id: &relationship_id,
            reltype: slide_type,
            target: &target,
            mode: TargetMode::Internal,
        }],
        &[],
        policy,
    )?;
    let shape_limits = source
        .owner_state
        .semantic_limits
        .as_deref()
        .copied()
        .unwrap_or_default()
        .shape();
    let placeholders =
        super::new_slide::plan(source.package.get_part(&layout)?.blob(), shape_limits)?;
    let mut slide_xml = super::new_slide::slide_xml_with_limits(id, &placeholders, shape_limits)?;
    if strict {
        for (old, new) in [
            (
                "http://schemas.openxmlformats.org/presentationml/2006/main",
                "http://purl.oclc.org/ooxml/presentationml/main",
            ),
            (
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://purl.oclc.org/ooxml/drawingml/main",
            ),
            (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
            ),
        ] {
            slide_xml = slide_xml.replace(old, new);
        }
    }
    let mut slide = BlobPart::new(
        name.clone(),
        ct::PML_SLIDE.to_owned(),
        slide_xml.into_bytes(),
    );
    slide.rels_mut().try_add_relationship(
        layout_type.to_owned(),
        layout.relative_ref(name.base_uri()),
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    let xml =
        super::xml::insert_slide(main.blob(), &source.slides, position, id, &relationship_id)?;
    let mut order: Vec<u32> = source.slides.iter().map(|slide| slide.id).collect();
    order.insert(position, id);
    let xml = super::sections::place(&xml, &order, position)?.unwrap_or(xml);
    super::new_slide::read_back_with_limits(slide.blob(), &placeholders, shape_limits)?;
    let sizes = [
        main.blob().len(),
        xml.len(),
        slide.blob().len(),
        types.bytes().len(),
        types_plan.final_len(),
        rels.bytes().len(),
        rels_plan.final_len(),
        slide.rels().plan_canonical(policy)?.final_len(),
    ];
    let retained = sizes
        .into_iter()
        .try_fold(0usize, |total, size| total.checked_add(size))
        .ok_or_else(|| invalid("slide insertion byte accounting overflow"))?;
    if retained > maximum {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::RetainedBytes),
            observed: retained,
            maximum,
        });
    }
    let next_types = types_plan.materialize(policy)?;
    let next_rels = rels_plan.materialize(policy)?;
    let mut candidate = source.package.as_ref().clone();
    candidate.try_add_parts_with_source_tokens(
        types.bytes(),
        &next_types,
        &rels,
        &next_rels,
        vec![Box::new(slide)],
    )?;
    candidate
        .get_part_mut(&source.presentation_name)?
        .set_blob(xml);
    let captured = super::model::capture_with_provenance_and_semantic_limits(
        &candidate,
        source.limits,
        source.physical_source_provenance,
        source.owner_state.semantic_limits.clone(),
    )?;
    if captured
        .slides
        .get(position)
        .is_none_or(|slide| slide.id != id || slide.part_name != name)
    {
        return Err(invalid("inserted slide identity failed native readback"));
    }
    Patch::capture(
        &source.package,
        &candidate,
        source.presentation_name.clone(),
        source.limits,
    )
}

// Existing physical names and dangling references both reserve identities:
// creating a previously missing target must not activate an unrelated edge.
fn available_name(source: &Snapshot) -> Result<PackURI> {
    let mut used = std::collections::HashSet::new();
    let mut reserve = |name: &str| -> Result<()> {
        const PREFIX: &str = "/ppt/slides/slide";
        if !name
            .get(..PREFIX.len())
            .is_some_and(|p| p.eq_ignore_ascii_case(PREFIX))
            || !name
                .get(name.len().saturating_sub(4)..)
                .is_some_and(|p| p.eq_ignore_ascii_case(".xml"))
        {
            return Ok(());
        }
        let Some(number) = name
            .get(PREFIX.len()..name.len().saturating_sub(4))
            .and_then(|s| s.parse::<usize>().ok())
        else {
            return Ok(());
        };
        if used.contains(&number) {
            return Ok(());
        }
        if used.len() >= source.limits.max_parts() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: used.len().checked_add(1).unwrap_or(usize::MAX),
                maximum: source.limits.max_parts(),
            });
        }
        used.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "slide insertion reserved names",
            source,
        })?;
        used.insert(number);
        Ok(())
    };
    for part in source.package.iter_parts() {
        reserve(part.partname().as_str())?;
    }
    for graph in
        std::iter::once(source.package.rels()).chain(source.package.iter_parts().map(|p| p.rels()))
    {
        for edge in graph.iter().filter(|edge| !edge.is_external()) {
            reserve(edge.target_partname()?.as_str())?;
        }
    }
    let maximum = source
        .package
        .part_count()
        .checked_add(used.len())
        .and_then(|n| n.checked_add(1))
        .ok_or_else(|| invalid("slide insertion name bound overflow"))?;
    for index in 1..=maximum {
        if used.contains(&index) {
            continue;
        }
        let name = PackURI::new(format!("/ppt/slides/slide{index}.xml")).map_err(Error::Invalid)?;
        if source.package.validate_new_part_name(&name).is_ok() {
            return Ok(name);
        }
    }
    Err(refuse("no collision-free slide part name is available"))
}
