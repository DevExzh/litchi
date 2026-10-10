//! Strict dependency-free admission against an effective structural catalog.
use super::Package;
use crate::presentation::order::StableSlideKey;
use crate::{Error, Result, SlideRemovalRefusal};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{PackURI, PreparedTopology, Relationships};

pub(super) fn validate(
    package: &Package,
    prepared: &PreparedTopology,
    root: &[u8],
    key: &StableSlideKey,
) -> Result<()> {
    validate_inner(package, prepared, root, key, false)
}

/// Validate an added slide while leaving unrelated opaque/dangling edges in
/// the source graph untouched. Their target identities are still reserved by
/// insertion planning; only an edge that actually targets the new slide can
/// make the candidate ambiguous. Direct removal uses the stricter validator
/// above and continues to reject such source-wide edges.
pub(super) fn validate_insertion(
    package: &Package,
    prepared: &PreparedTopology,
    root: &[u8],
    key: &StableSlideKey,
) -> Result<()> {
    validate_inner(package, prepared, root, key, true)
}

fn validate_inner(
    package: &Package,
    prepared: &PreparedTopology,
    root: &[u8],
    key: &StableSlideKey,
    allow_unrelated_opaque_targets: bool,
) -> Result<()> {
    prepared.check_source()?;
    prepared.check_execution()?;
    let source = &package.owner.package;
    let presentation_uri = package.owner._presentation.source_partname();
    crate::presentation::source_removal::validate_primary_root(
        source,
        presentation_uri,
        root,
        &key.relationship_id,
    )?;
    let presentation = prepared.part(presentation_uri)?;
    if presentation.content_type() != ct::PML_PRESENTATION_MAIN {
        return refuse(
            SlideRemovalRefusal::AmbiguousTopology,
            "unexpected presentation content type",
        );
    }
    let selected = presentation
        .rels()
        .get(&key.relationship_id)
        .ok_or(Error::StaleSource)?;
    if selected.is_external()
        || !crate::parts::is_relationship_type(selected.reltype(), rt::SLIDE, "slide")
        || selected.target_query().is_some()
        || selected.target_fragment().is_some()
        || !selected.target_partname()?.is_equivalent_to(&key.part_uri)
    {
        return refuse(
            SlideRemovalRefusal::AmbiguousTopology,
            "effective slide binding differs from catalog",
        );
    }
    let slide = prepared.part(&key.part_uri)?;
    if slide.content_type() != ct::PML_SLIDE || slide.rels().len() != 1 {
        return refuse(
            SlideRemovalRefusal::UnsupportedRelationship,
            "selected slide must have one layout relationship",
        );
    }
    let layout = slide.rels().iter().next().ok_or(Error::StaleSource)?;
    if layout.is_external()
        || !crate::parts::is_relationship_type(layout.reltype(), rt::SLIDE_LAYOUT, "slideLayout")
        || layout.target_query().is_some()
        || layout.target_fragment().is_some()
    {
        return refuse(
            SlideRemovalRefusal::UnsupportedRelationship,
            "selected slide has a non-layout relationship",
        );
    }
    canonical_edge(presentation_uri, selected)?;
    canonical_edge(&key.part_uri, layout)?;
    let layout_uri = layout.target_partname()?;
    if !layout_uri.as_str().starts_with("/ppt/slideLayouts/")
        || prepared.part(&layout_uri)?.content_type() != ct::PML_SLIDE_LAYOUT
    {
        return refuse(
            SlideRemovalRefusal::AmbiguousTopology,
            "selected layout has no native layout owner",
        );
    }
    // Structural recipes change only slide Parts and presentation bindings;
    // the master's registered layout graph remains the physical source graph.
    crate::presentation::source_cross_copy::validate_registered_source_layout(
        source,
        &source.part(presentation_uri)?,
        package.owner._presentation.source_data().as_bytes(),
        &layout_uri,
        source.execution_context().as_ref(),
    )?;
    let data = slide.data()?;
    crate::source_limits::validate_xml_limits(data.as_bytes(), package.owner.limits)?;
    crate::opened::validate_removal_slide_surface(data.as_bytes())?;

    let maximum = package
        .owner
        .limits
        .max_parts()
        .checked_mul(16)
        .ok_or_else(|| Error::Invalid("slide owner scan limit overflow".into()))?;
    let mut scanned = 0usize;
    let mut incoming = 0usize;
    let mut visit = |owner: Option<&PackURI>, relationships: &Relationships| -> Result<()> {
        for relationship in relationships.iter() {
            prepared.check_execution()?;
            scanned = scanned
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("slide owner scan overflow".into()))?;
            if scanned > maximum {
                return Err(Error::Limit {
                    resource: "source slide removal relationships",
                    limit: maximum,
                });
            }
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            if !target.is_equivalent_to(&key.part_uri) {
                if allow_unrelated_opaque_targets {
                    continue;
                }
                if relationship.target_query().is_some() || relationship.target_fragment().is_some()
                {
                    return refuse(
                        SlideRemovalRefusal::AmbiguousTopology,
                        "internal relationship has a query or fragment",
                    );
                }
                continue;
            }
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return refuse(
                    SlideRemovalRefusal::AmbiguousTopology,
                    "internal relationship has a query or fragment",
                );
            }
            if owner.is_some_and(|owner| owner.is_equivalent_to(presentation_uri))
                && relationship.r_id() == key.relationship_id
                && crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
            {
                incoming += 1;
            } else {
                return refuse(
                    SlideRemovalRefusal::SharedOwner,
                    "another owner references selected slide",
                );
            }
        }
        Ok(())
    };
    visit(None, prepared.rels())?;
    for owner in prepared.iter_parts()? {
        visit(Some(owner.partname()), owner.rels())?;
    }
    if incoming != 1 {
        return refuse(
            SlideRemovalRefusal::AmbiguousTopology,
            "selected slide must have one presentation owner",
        );
    }
    prepared.check_source()?;
    prepared.check_execution()?;
    Ok(())
}

fn refuse<T>(kind: SlideRemovalRefusal, detail: &str) -> Result<T> {
    Err(Error::SlideRemovalPlan {
        kind,
        detail: detail.to_owned(),
    })
}

/// Add the single typed layout edge before installing its exact XML image.
/// OPC's raw-image guard subsequently validates the complete manifest grammar
/// and proves that the image describes this typed edge and no others.
pub(super) fn stage_layout(
    package: &Package,
    plan: &mut litchi_opc::SourceTopologyPlan,
    owner: &PackURI,
    xml: &[u8],
) -> Result<()> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::events::Event;
    use quick_xml::name::{Namespace, ResolveResult};
    use quick_xml::reader::NsReader;
    crate::source_limits::validate_xml_limits(xml, package.owner.limits)?;
    let amount = xml
        .len()
        .checked_mul(4)
        .and_then(|n| n.checked_add(1024))
        .ok_or_else(|| Error::Invalid("restored relationship workspace overflow".into()))?;
    let context = package.owner.package.execution_context();
    let _reservation = context
        .as_ref()
        .map(|context| {
            context
                .reserve(litchi_core::Resource::Memory, amount as u64)
                .map_err(|error| {
                    Error::Opc(match error {
                        litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                        other => litchi_opc::OpcError::Execution(other),
                    })
                })
        })
        .transpose()?;
    let mut reader = NsReader::from_reader(xml);
    let mut edge = None;
    loop {
        package.owner.package.check_execution()?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|e| Error::Xml(e.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship"
                    && matches!(
                        namespace,
                        ResolveResult::Bound(Namespace(
                            b"http://schemas.openxmlformats.org/package/2006/relationships"
                        ))
                    ) =>
            {
                if edge.is_some() {
                    return refuse(
                        SlideRemovalRefusal::UnsupportedRelationship,
                        "restored slide must have one layout edge",
                    );
                }
                let mut id = None;
                let mut kind = None;
                let mut target = None;
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|e| Error::Xml(e.to_string()))?;
                    let value = attribute
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    match attribute.key.as_ref() {
                        b"Id" => id = Some(value.into_owned()),
                        b"Type" => kind = Some(value.into_owned()),
                        b"Target" => target = Some(value.into_owned()),
                        b"TargetMode" if value.as_ref() != "Internal" => {
                            return refuse(
                                SlideRemovalRefusal::UnsupportedRelationship,
                                "restored layout edge must be internal",
                            );
                        },
                        _ => {},
                    }
                }
                let id = id.ok_or(Error::StaleSource)?;
                let kind = kind.ok_or(Error::StaleSource)?;
                let target = target.ok_or(Error::StaleSource)?;
                if !crate::parts::is_relationship_type(&kind, rt::SLIDE_LAYOUT, "slideLayout")
                    || target.contains(['?', '#'])
                {
                    return refuse(
                        SlideRemovalRefusal::UnsupportedRelationship,
                        "restored slide has unsupported edge",
                    );
                }
                let target =
                    PackURI::from_rel_ref(owner.base_uri(), &target).map_err(Error::Invalid)?;
                edge = Some((id, kind, target));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let (id, kind, target) = edge.ok_or(Error::StaleSource)?;
    plan.try_add_internal_relationship(owner.clone(), id, kind, target)?;
    Ok(())
}

/// Typed OPC additions use owner-relative references. Until their raw-image
/// restoration accepts equivalent alternate reference spellings, reject a
/// forward removal that could not be inverted after reopening.
pub(super) fn validate_physical_roundtrip(package: &Package, key: &StableSlideKey) -> Result<()> {
    let source = &package.owner.package;
    let presentation_uri = package.owner._presentation.source_partname();
    let presentation = source.part(presentation_uri)?;
    canonical_edge(
        presentation_uri,
        presentation
            .rels()
            .get(&key.relationship_id)
            .ok_or(Error::StaleSource)?,
    )?;
    let slide = source.part(&key.part_uri)?;
    let layout = slide.rels().iter().next().ok_or(Error::StaleSource)?;
    canonical_edge(&key.part_uri, layout)
}

fn canonical_edge(owner: &PackURI, relationship: &litchi_opc::Relationship) -> Result<()> {
    if relationship.target_ref()
        != relationship
            .target_partname()?
            .relative_ref(owner.base_uri())
    {
        return refuse(
            SlideRemovalRefusal::AmbiguousTopology,
            "primary durable removal requires canonical owner-relative relationship targets",
        );
    }
    Ok(())
}
