//! Package-level authoring and relationship-graph operations.

use crate::parts::{SlideLayoutPart, SlideMasterPart};
use crate::{Error, Result};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashSet;

use super::codec::{
    IdListAnchor, MAX_NAME_CHARS, MAX_PLACEHOLDERS_PER_OPERATION, MAX_SCAN_DEPTH, MAX_SCAN_NODES,
    P_NS, R_NS, SPTREE_DEPTH, STRICT_SLIDE_LAYOUT_REL, STRICT_SLIDE_MASTER_REL, check_size,
    escape_xml, find_placeholder_span, insert_bytes, insert_id_list_entry, invalid, layout_xml,
    local_name, master_xml, next_shape_id, placeholder_shape_xml, remove_id_list_entry,
    replace_span, scan_element_span, shape_id_within,
};
use super::model::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderSpec,
    SlideLayoutKind,
};

// ============================================================================
// Authoring operations
// ============================================================================

/// Create a new slide master and reference it from the presentation part.
///
/// The master is written with a color map, an empty `p:sldLayoutIdLst`, and
/// default `p:txStyles` (title, body, and other styles with nine paragraph
/// levels each). It is related to an existing theme part when one exists,
/// otherwise a new theme part is generated. The presentation part gains a
/// slide-master relationship plus a `p:sldMasterId` entry whose ID is one
/// above the current maximum (starting at [`MIN_MASTER_OR_LAYOUT_ID`]).
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn add_slide_master(package: &mut OpcPackage) -> Result<AuthoredSlideMaster> {
    let presentation_name = package.main_document_part()?.partname().clone();
    require_presentation_part(package.get_part(&presentation_name)?.content_type())?;
    let presentation_xml = package.get_part(&presentation_name)?.blob().to_vec();
    let entries = parse_master_id_list(&presentation_xml)?;
    let master_id = allocate_id(entries.iter().map(|entry| entry.0))?;

    let master_index = next_part_index(package, "/ppt/slideMasters/slideMaster", ".xml")?;
    let master_uri = PackURI::new(format!("/ppt/slideMasters/slideMaster{master_index}.xml"))
        .map_err(|error| Error::Uri(format!("slide master partname: {error}")))?;
    let master_dir = "/ppt/slideMasters/";

    let (theme_target, new_theme) = theme_target_for_new_master(package, master_dir)?;

    let presentation = package.get_part_mut(&presentation_name)?;
    let relationship_id = presentation.relate_to(
        &format!("slideMasters/slideMaster{master_index}.xml"),
        rt::SLIDE_MASTER,
    );
    let entry = format!(
        "<p:sldMasterId xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\" id=\"{master_id}\" r:id=\"{}\"/>",
        escape_xml(&relationship_id)
    );
    let patched = match insert_id_list_entry(
        &presentation_xml,
        "sldMasterIdLst",
        &entry,
        IdListAnchor::AfterRootStart,
    ) {
        Ok(patched) => patched,
        Err(error) => {
            presentation.rels_mut().remove(&relationship_id);
            return Err(error);
        },
    };
    presentation.set_blob(patched);

    if let Some((theme_uri, theme_xml)) = new_theme {
        package.add_part(Box::new(BlobPart::new(
            theme_uri,
            ct::OFC_THEME.to_string(),
            theme_xml.into_bytes(),
        )));
    }
    let mut master_part = BlobPart::new(
        master_uri.clone(),
        ct::PML_SLIDE_MASTER.to_string(),
        master_xml().into_bytes(),
    );
    master_part.relate_to(&theme_target, rt::THEME);
    package.add_part(Box::new(master_part));

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(AuthoredSlideMaster {
        master_id,
        relationship_id,
        part_name: master_uri.clone(),
    })
}

/// Create a new slide layout attached to an existing slide master.
///
/// The master gains a slide-layout relationship plus a `p:sldLayoutId` entry
/// whose ID is one above the current maximum within that master. The layout
/// part is written with the given `ST_SlideLayoutType` kind, name, and
/// optional placeholder shapes, and carries the required relationship back to
/// its owning master.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn add_slide_layout(
    package: &mut OpcPackage,
    master_part_name: &PackURI,
    kind: SlideLayoutKind,
    name: &str,
    placeholders: &[PlaceholderSpec],
) -> Result<AuthoredSlideLayout> {
    require_name(name)?;
    require_placeholders(placeholders)?;
    let master_uri = master_part_name.clone();
    let master_part = package.get_part(&master_uri)?;
    if master_part.content_type() != ct::PML_SLIDE_MASTER {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE_MASTER.to_string(),
            actual: master_part.content_type().to_string(),
        });
    }
    let references = parse_layout_id_list(master_part.blob())?;
    let layout_id = allocate_id(references.iter().filter_map(LayoutReference::layout_id))?;
    let master_xml = master_part.blob().to_vec();

    let layout_index = next_part_index(package, "/ppt/slideLayouts/slideLayout", ".xml")?;
    let layout_uri = PackURI::new(format!("/ppt/slideLayouts/slideLayout{layout_index}.xml"))
        .map_err(|error| Error::Uri(format!("slide layout partname: {error}")))?;
    let layout_xml = layout_xml(kind, name, placeholders)?;

    let master_part = package.get_part_mut(&master_uri)?;
    let relationship_id = master_part.relate_to(
        &format!("../slideLayouts/slideLayout{layout_index}.xml"),
        rt::SLIDE_LAYOUT,
    );
    let entry = format!(
        "<p:sldLayoutId xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\" id=\"{layout_id}\" r:id=\"{}\"/>",
        escape_xml(&relationship_id)
    );
    let patched = match insert_id_list_entry(
        &master_xml,
        "sldLayoutIdLst",
        &entry,
        IdListAnchor::AfterElement("clrMap"),
    ) {
        Ok(patched) => patched,
        Err(error) => {
            master_part.rels_mut().remove(&relationship_id);
            return Err(error);
        },
    };
    master_part.set_blob(patched);

    let mut layout_part = BlobPart::new(
        layout_uri.clone(),
        ct::PML_SLIDE_LAYOUT.to_string(),
        layout_xml.into_bytes(),
    );
    layout_part.relate_to(
        &relative_target("/ppt/slideLayouts/", master_part_name.as_str())?,
        rt::SLIDE_MASTER,
    );
    package.add_part(Box::new(layout_part));

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(AuthoredSlideLayout {
        layout_id,
        relationship_id,
        part_name: layout_uri.clone(),
        master_part_name: master_uri.clone(),
    })
}

/// Add or replace a placeholder shape on a slide master or slide layout.
///
/// A placeholder is identified by its `p:ph` type and index (an absent index
/// matches `idx` zero, per the ECMA default). When a matching placeholder
/// already exists its shape is replaced in place, keeping its shape ID;
/// otherwise a new shape is appended to the shape tree with the next free
/// shape ID. The optional text replaces the placeholder's prompt text.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store_placeholder_shape(
    package: &mut OpcPackage,
    part_name: &PackURI,
    spec: &PlaceholderSpec,
) -> Result<()> {
    let uri = part_name.clone();
    let part = package.get_part(&uri)?;
    let content_type = part.content_type();
    if content_type != ct::PML_SLIDE_MASTER && content_type != ct::PML_SLIDE_LAYOUT {
        return Err(Error::ContentType {
            expected: format!("{} or {}", ct::PML_SLIDE_MASTER, ct::PML_SLIDE_LAYOUT),
            actual: content_type.to_string(),
        });
    }
    if let Some(name) = &spec.name {
        require_name(name)?;
    }
    let xml = part.blob().to_vec();
    let existing = find_placeholder_span(&xml, spec.kind.as_str(), spec.effective_index())?;
    let shape_id = match &existing {
        Some(span) => shape_id_within(&xml[span.start..span.end])?,
        None => next_shape_id(&xml)?,
    };
    let shape = placeholder_shape_xml(shape_id, spec, true);
    let patched = if let Some(span) = existing {
        replace_span(&xml, &span, shape.as_bytes())?
    } else {
        let tree = scan_element_span(&xml, "spTree", SPTREE_DEPTH)?
            .ok_or_else(|| invalid("slide master or layout has no shape tree"))?;
        if tree.empty {
            return Err(invalid("slide master or layout has an empty shape tree"));
        }
        insert_bytes(&xml, tree.close_start, shape.as_bytes())?
    };
    // The patched part must inventory the placeholder back through the same
    // scan the read side's shape parser performs.
    if find_placeholder_span(&patched, spec.kind.as_str(), spec.effective_index())?.is_none() {
        return Err(invalid("patched placeholder shape did not round-trip"));
    }
    package.get_part_mut(&uri)?.set_blob(patched);

    // Run the read-side placeholder inventory over the patched part.
    let part = package.get_part(&uri)?;
    let matches = |shape: crate::shape::Shape<'_>| {
        shape.placeholder().is_some_and(|placeholder| {
            placeholder.kind().unwrap_or("obj") == spec.kind.as_str()
                && placeholder.index() == spec.effective_index()
        })
    };
    let found = if part.content_type() == ct::PML_SLIDE_MASTER {
        SlideMasterPart::from_part(part)?
            .shapes()?
            .placeholders()
            .any(matches)
    } else {
        SlideLayoutPart::from_part(part)?
            .shapes()?
            .placeholders()
            .any(matches)
    };
    if !found {
        return Err(invalid(
            "read-side placeholder inventory lost the authored shape",
        ));
    }
    invalidate_signatures(package);
    Ok(())
}

/// Delete a slide layout that is not referenced by any slide.
///
/// The owning master's `p:sldLayoutIdLst` entry and relationship are removed
/// together with the layout part itself. Layouts still referenced by a slide,
/// or not owned by any master, are rejected.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn remove_slide_layout(package: &mut OpcPackage, layout_part_name: &PackURI) -> Result<()> {
    let layout_uri = layout_part_name.clone();
    let layout_part = package.get_part(&layout_uri)?;
    if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE_LAYOUT.to_string(),
            actual: layout_part.content_type().to_string(),
        });
    }

    // Reject layouts still referenced by a slide.
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE {
            continue;
        }
        for relationship in part.rels().iter() {
            if matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_REL
            ) && !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target == layout_uri)
            {
                return Err(invalid(format!(
                    "slide layout '{layout_part_name}' is still referenced by slide '{}'",
                    part.partname()
                )));
            }
        }
    }

    // Locate the owning master entry.
    let mut owner = None;
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE_MASTER {
            continue;
        }
        for reference in parse_layout_id_list(part.blob())? {
            let Some(relationship) = part.rels().get(reference.relationship_id()) else {
                continue;
            };
            if !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target == layout_uri)
            {
                if owner.is_some() {
                    return Err(invalid(format!(
                        "slide layout '{layout_part_name}' is owned by more than one master"
                    )));
                }
                owner = Some((
                    part.partname().clone(),
                    reference.relationship_id().to_string(),
                ));
            }
        }
    }
    let (master_uri, relationship_id) = owner.ok_or_else(|| {
        invalid(format!(
            "slide layout '{layout_part_name}' is not referenced by any slide master"
        ))
    })?;

    let master_xml = package.get_part(&master_uri)?.blob().to_vec();
    let patched = remove_id_list_entry(&master_xml, "sldLayoutId", &relationship_id)?;
    let master_part = package.get_part_mut(&master_uri)?;
    master_part.set_blob(patched);
    if master_part.rels_mut().remove(&relationship_id).is_none() {
        return Err(Error::Relationship(format!(
            "slide master lost slide-layout relationship '{relationship_id}'"
        )));
    }
    package.remove_part(&layout_uri);

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(())
}

// ============================================================================
// Graph validation
// ============================================================================

/// Validate the slide master and slide layout graph of a package.
///
/// This mirrors the rules the read side applies when resolving
/// `Presentation::slide_masters`, `SlideMaster::slide_layouts`, and
/// `SlideLayout::master`:
///
/// - every `p:sldMasterId` entry has a unique ID ≥ [`MIN_MASTER_OR_LAYOUT_ID`]
///   and resolves through an internal slide-master relationship to a part
///   with the slide-master content type;
/// - every `p:sldLayoutId` entry of each master resolves through an internal
///   slide-layout relationship to a part with the slide-layout content type;
/// - every referenced layout has exactly one internal slide-master
///   relationship, pointing back to the master that references it.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn validate_master_layout_graph(package: &OpcPackage) -> Result<()> {
    let presentation = package.main_document_part()?;
    require_presentation_part(presentation.content_type())?;
    let entries = parse_master_id_list(presentation.blob())?;
    let mut master_parts = HashSet::new();
    for (master_id, relationship_id) in &entries {
        let relationship = presentation.rels().get(relationship_id).ok_or_else(|| {
            Error::Relationship(format!(
                "slide master ID {master_id} references missing relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::SLIDE_MASTER | STRICT_SLIDE_MASTER_REL
            )
        {
            return Err(Error::Relationship(format!(
                "relationship '{relationship_id}' is not an internal slide-master relationship"
            )));
        }
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!(
                "invalid slide-master relationship '{relationship_id}': {error}"
            ))
        })?;
        let master_part = package.get_part(&target)?;
        if master_part.content_type() != ct::PML_SLIDE_MASTER {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE_MASTER.to_string(),
                actual: master_part.content_type().to_string(),
            });
        }
        if !master_parts.insert(target.to_string()) {
            return Err(invalid(format!(
                "slide master part '{target}' is referenced more than once"
            )));
        }
        validate_master_layouts(package, master_part)?;
    }
    Ok(())
}

fn validate_master_layouts(package: &OpcPackage, master_part: &dyn Part) -> Result<()> {
    let references = parse_layout_id_list(master_part.blob())?;
    for reference in &references {
        let relationship_id = reference.relationship_id();
        let relationship = master_part.rels().get(relationship_id).ok_or_else(|| {
            Error::Relationship(format!(
                "slide master references missing slide-layout relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_REL
            )
        {
            return Err(Error::Relationship(format!(
                "relationship '{relationship_id}' is not an internal slide-layout relationship"
            )));
        }
        let layout_name = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!(
                "invalid slide-layout relationship '{relationship_id}': {error}"
            ))
        })?;
        let layout_part = package.get_part(&layout_name)?;
        if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE_LAYOUT.to_string(),
                actual: layout_part.content_type().to_string(),
            });
        }
        // The layout must keep exactly one internal relationship back to the
        // master that references it.
        let mut back_reference = None;
        for candidate in layout_part.rels().iter() {
            if matches!(
                candidate.reltype(),
                rt::SLIDE_MASTER | STRICT_SLIDE_MASTER_REL
            ) {
                if back_reference.is_some() {
                    return Err(Error::Relationship(format!(
                        "slide layout '{layout_name}' has multiple slide-master relationships"
                    )));
                }
                back_reference = Some(candidate);
            }
        }
        let back_reference = back_reference.ok_or_else(|| {
            Error::Relationship(format!(
                "slide layout '{layout_name}' has no slide-master relationship"
            ))
        })?;
        if back_reference.is_external()
            || back_reference
                .target_partname()
                .is_ok_and(|target| target != *master_part.partname())
        {
            return Err(Error::Relationship(format!(
                "slide layout '{layout_name}' does not reference its owning master '{}'",
                master_part.partname()
            )));
        }
    }
    Ok(())
}

// ============================================================================
// Presentation master-ID parsing and allocation
// ============================================================================

#[derive(Debug, Clone)]
struct LayoutReference {
    layout_id: Option<u32>,
    relationship_id: String,
}

impl LayoutReference {
    fn layout_id(&self) -> Option<u32> {
        self.layout_id
    }

    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
}

/// Parse the layout ID entries owned by one slide master.
fn parse_layout_id_list(xml: &[u8]) -> Result<Vec<LayoutReference>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut in_list = false;
    let mut entries = Vec::new();
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML resource limit exceeded"))?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML nesting is too deep"))?;
                if nodes > MAX_SCAN_NODES || depth > MAX_SCAN_DEPTH {
                    return Err(invalid("slide-master XML resource limit exceeded"));
                }
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                if depth == 2 && local == b"sldLayoutIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    in_list = true;
                } else if depth == 3 && in_list && local == b"sldLayoutId" {
                    push_layout_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::Empty(element)) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("slide-master XML resource limit exceeded"))?;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("slide-master XML resource limit exceeded"));
                }
                let element_name = element.name();
                let local = local_name(element_name.as_ref());
                if depth == 1 && local == b"sldLayoutIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-layout ID list"));
                    }
                    in_list = true;
                } else if depth == 2 && in_list && local == b"sldLayoutId" {
                    push_layout_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::End(element)) => {
                if depth == 0 {
                    return Err(invalid("unexpected closing element in slide-master XML"));
                }
                if depth == 2 && local_name(element.name().as_ref()) == b"sldLayoutIdLst" {
                    in_list = false;
                }
                depth -= 1;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated slide-master XML"));
    }
    Ok(entries)
}

fn push_layout_id_entry(
    entries: &mut Vec<LayoutReference>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    let mut layout_id = None;
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if name == "id" {
            let parsed = value
                .parse::<u32>()
                .map_err(|_err| invalid(format!("invalid slide-layout ID '{value}'")))?;
            if parsed < MIN_MASTER_OR_LAYOUT_ID {
                return Err(invalid(format!(
                    "slide-layout ID {parsed} is below {MIN_MASTER_OR_LAYOUT_ID}"
                )));
            }
            layout_id = Some(parsed);
        } else if name.rsplit_once(':').map(|(_, local)| local) == Some("id") {
            relationship_id = Some(value.to_owned());
        }
    }
    let relationship_id = relationship_id
        .ok_or_else(|| invalid("slide-layout entry is missing its relationship ID"))?;
    if relationship_id.is_empty() {
        return Err(invalid("empty slide-layout relationship ID"));
    }
    if entries
        .iter()
        .any(|entry| entry.layout_id == layout_id && layout_id.is_some())
    {
        return Err(invalid("duplicate slide-layout ID"));
    }
    if entries
        .iter()
        .any(|entry| entry.relationship_id == relationship_id)
    {
        return Err(invalid(format!(
            "duplicate slide-layout relationship ID '{relationship_id}'"
        )));
    }
    entries.push(LayoutReference {
        layout_id,
        relationship_id,
    });
    Ok(())
}

/// Parse `p:sldMasterIdLst` entries as `(id, relationship_id)` pairs.
///
/// Mirrors the read-side rules: IDs are unsigned 32-bit values at or above
/// [`MIN_MASTER_OR_LAYOUT_ID`], and both IDs and relationship IDs are unique.
fn parse_master_id_list(xml: &[u8]) -> Result<Vec<(u32, String)>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut in_list = false;
    let mut entries = Vec::new();
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                depth += 1;
                if nodes > MAX_SCAN_NODES || depth > MAX_SCAN_DEPTH {
                    return Err(invalid("presentation XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == 2 && local == b"sldMasterIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-master ID list"));
                    }
                    in_list = true;
                } else if depth == 3 && in_list && local == b"sldMasterId" {
                    push_master_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("presentation XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == 1 && local == b"sldMasterIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-master ID list"));
                    }
                    in_list = true;
                } else if depth == 2 && in_list && local == b"sldMasterId" {
                    push_master_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::End(element)) => {
                if depth == 2 && local_name(element.name().as_ref()) == b"sldMasterIdLst" {
                    in_list = false;
                }
                if depth == 0 {
                    return Err(invalid("unexpected closing element in presentation XML"));
                }
                depth -= 1;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated presentation XML"));
    }
    Ok(entries)
}

fn push_master_id_entry(
    entries: &mut Vec<(u32, String)>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    let mut id = None;
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if name == "id" {
            let parsed = value
                .parse::<u32>()
                .map_err(|_err| invalid(format!("invalid slide-master ID '{value}'")))?;
            if parsed < MIN_MASTER_OR_LAYOUT_ID {
                return Err(invalid(format!(
                    "slide-master ID {parsed} is below {MIN_MASTER_OR_LAYOUT_ID}"
                )));
            }
            id = Some(parsed);
        } else if name.rsplit_once(':').map(|(_, local)| local) == Some("id") {
            relationship_id = Some(value.to_owned());
        }
    }
    let id = id.ok_or_else(|| invalid("slide-master entry is missing its ID"))?;
    let relationship_id = relationship_id
        .ok_or_else(|| invalid("slide-master entry is missing its relationship ID"))?;
    if relationship_id.is_empty() {
        return Err(invalid("empty slide-master relationship ID"));
    }
    if entries.iter().any(|(existing, _)| *existing == id) {
        return Err(invalid(format!("duplicate slide-master ID {id}")));
    }
    if entries
        .iter()
        .any(|(_, existing)| *existing == relationship_id)
    {
        return Err(invalid(format!(
            "duplicate slide-master relationship ID '{relationship_id}'"
        )));
    }
    entries.push((id, relationship_id));
    Ok(())
}

/// Allocate the next ID above the current maximum.
fn allocate_id(used: impl Iterator<Item = u32>) -> Result<u32> {
    used.max()
        .unwrap_or(MIN_MASTER_OR_LAYOUT_ID - 1)
        .checked_add(1)
        .ok_or_else(|| invalid("slide master or layout ID overflow"))
}

/// Find the lowest free numeric suffix for a part-name pattern.
fn next_part_index(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<u32> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| Error::Uri(format!("partname allocation: {error}")))?;
        if package.get_part(&candidate).is_err() {
            return Ok(index);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("part-name index overflow"))?;
    }
}

/// Resolve the theme relationship target for a newly created master.
///
/// Returns the relationship target (relative to `/ppt/slideMasters/`) and,
/// when no theme exists yet, a new theme part to add.
fn theme_target_for_new_master(
    package: &OpcPackage,
    master_dir: &str,
) -> Result<(String, Option<(PackURI, String)>)> {
    // Prefer the theme used by an existing slide master.
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE_MASTER {
            continue;
        }
        for relationship in part.rels().iter() {
            if relationship.reltype() == rt::THEME
                && !relationship.is_external()
                && let Ok(target) = relationship.target_partname()
                && package.get_part(&target).is_ok()
            {
                return Ok((relative_target(master_dir, target.as_str())?, None));
            }
        }
    }
    // Otherwise reuse any existing theme part.
    for part in package.iter_parts() {
        if part.content_type() == ct::OFC_THEME {
            return Ok((relative_target(master_dir, part.partname().as_str())?, None));
        }
    }
    // Otherwise author a fresh theme part from the default template.
    let index = next_part_index(package, "/ppt/theme/theme", ".xml")?;
    let uri = PackURI::new(format!("/ppt/theme/theme{index}.xml"))
        .map_err(|error| Error::Uri(format!("theme partname: {error}")))?;
    Ok((
        format!("../theme/theme{index}.xml"),
        Some((uri, crate::resources::THEME.to_string())),
    ))
}

/// Compute the relationship target for `target` relative to `source_dir`.
///
/// Both names must be absolute pack URIs; the result uses `..` segments to
/// climb out of the source directory.
fn relative_target(source_dir: &str, target: &str) -> Result<String> {
    let source = source_dir.trim_matches('/');
    let target = target.trim_start_matches('/');
    let source_segments: Vec<&str> = source.split('/').filter(|item| !item.is_empty()).collect();
    let target_segments: Vec<&str> = target.split('/').filter(|item| !item.is_empty()).collect();
    let common = source_segments
        .iter()
        .zip(target_segments.iter())
        .take_while(|(left, right)| left == right)
        .count();
    if common == 0 && !source_segments.is_empty() {
        return Err(Error::Uri(format!(
            "cannot relativize '{target}' against '/{source}/'"
        )));
    }
    let mut result = String::new();
    for _ in common..source_segments.len() {
        result.push_str("../");
    }
    result.push_str(&target_segments[common..].join("/"));
    Ok(result)
}

// ============================================================================
// Misc validators
// ============================================================================

fn require_presentation_part(content_type: &str) -> Result<()> {
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

fn require_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid("slide layout name cannot be empty"));
    }
    if name.chars().count() > MAX_NAME_CHARS {
        return Err(invalid("slide layout name exceeds 256 characters"));
    }
    Ok(())
}

fn require_placeholders(placeholders: &[PlaceholderSpec]) -> Result<()> {
    if placeholders.len() > MAX_PLACEHOLDERS_PER_OPERATION {
        return Err(invalid("too many placeholder shapes in one operation"));
    }
    let mut identities = HashSet::new();
    for spec in placeholders {
        if let Some(name) = &spec.name {
            require_name(name)?;
        }
        if !identities.insert((spec.kind, spec.effective_index())) {
            return Err(invalid(format!(
                "duplicate placeholder type '{}' with index {}",
                spec.kind.as_str(),
                spec.effective_index()
            )));
        }
    }
    Ok(())
}

fn invalidate_signatures(package: &mut OpcPackage) {
    package.unsign();
}
