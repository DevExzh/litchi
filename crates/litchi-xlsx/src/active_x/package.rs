//! OPC package graph lifecycle for inert ActiveX resources.

use super::codec::{
    bounded, collect_binary_ids_descriptor, controls_span, relationship_ids_in_xml,
    replace_controls_xml, validate_controls, validate_descriptor,
};
use super::model::*;
use super::{
    BINARY_CONTENT_TYPE, BINARY_REL, CONTROL_REL, CONTROL_REL_STRICT, DESCRIPTOR_CONTENT_TYPE,
    IMAGE_REL, IMAGE_REL_STRICT, MAX_BINARY, MAX_CONTROLS, MAX_TOTAL_BINARY, Result,
    WORKSHEET_CONTENT_TYPE, content_type, invalid, limit, relerr,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use std::collections::HashSet;

/// Loads every ActiveX control referenced by one worksheet. All payload bytes remain inert.
pub fn load_from_worksheet(package: &OpcPackage, worksheet_uri: &PackURI) -> Result<ControlSet> {
    let worksheet = package.get_part(worksheet_uri)?;
    if worksheet.content_type() != WORKSHEET_CONTENT_TYPE {
        return Err(content_type(
            WORKSHEET_CONTENT_TYPE,
            worksheet.content_type(),
        ));
    }
    let parsed = Controls::parse(worksheet.blob())?;
    let referenced: HashSet<&str> = parsed
        .controls
        .iter()
        .map(|c| c.relationship_id.as_str())
        .collect();
    for rel in worksheet.rels().iter() {
        if matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT)
            && !referenced.contains(rel.r_id())
        {
            return Err(relerr(
                "worksheet has an unreferenced ActiveX control relationship",
            ));
        }
    }
    let mut loaded = Vec::with_capacity(parsed.controls.len());
    let mut total_binary = 0usize;
    for control in parsed.controls {
        let preview = if let Some(id) = control
            .properties
            .as_ref()
            .and_then(|p| p.preview_relationship_id.as_deref())
        {
            let rel = worksheet
                .rels()
                .get(id)
                .ok_or_else(|| relerr("control preview relationship is missing"))?;
            if rel.is_external() || !matches!(rel.reltype(), IMAGE_REL | IMAGE_REL_STRICT) {
                return Err(relerr(
                    "control preview must be an internal image relationship",
                ));
            }
            let part = package.get_part(&rel.target_partname()?)?;
            if !part.content_type().starts_with("image/") {
                return Err(content_type("image/*", part.content_type()));
            }
            if part.blob().len() > MAX_BINARY {
                return Err(limit("ActiveX preview image bytes"));
            }
            total_binary = total_binary
                .checked_add(part.blob().len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
            if total_binary > MAX_TOTAL_BINARY {
                return Err(limit("aggregate ActiveX resource bytes"));
            }
            Some(PreviewImage {
                relationship_id: id.to_string(),
                part_uri: rel.target_partname()?,
                content_type: part.content_type().to_string(),
                bytes: part.blob().to_vec(),
            })
        } else {
            None
        };
        let rel = worksheet
            .rels()
            .get(&control.relationship_id)
            .ok_or_else(|| relerr("control relationship is missing"))?;
        if rel.is_external() || !matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT) {
            return Err(relerr("control must target an internal ActiveX descriptor"));
        }
        let descriptor_uri = rel.target_partname()?;
        let part = package.get_part(&descriptor_uri)?;
        if part.content_type() != DESCRIPTOR_CONTENT_TYPE {
            return Err(content_type(DESCRIPTOR_CONTENT_TYPE, part.content_type()));
        }
        let descriptor = Descriptor::parse(part.blob())?;
        let mut ids = HashSet::new();
        collect_binary_ids_descriptor(&descriptor, &mut ids)?;
        if part.rels().iter().count() != ids.len() {
            return Err(relerr(
                "ActiveX descriptor has unexpected or duplicate outgoing relationships",
            ));
        }
        let mut binaries = Vec::with_capacity(ids.len());
        for id in ids {
            let binary_rel = part
                .rels()
                .get(&id)
                .ok_or_else(|| relerr("ActiveX binary relationship is missing"))?;
            if binary_rel.is_external() || binary_rel.reltype() != BINARY_REL {
                return Err(relerr(
                    "ActiveX descriptor may relate only to internal ActiveX binaries",
                ));
            }
            let binary_uri = binary_rel.target_partname()?;
            let binary = package.get_part(&binary_uri)?;
            if binary.content_type() != BINARY_CONTENT_TYPE {
                return Err(content_type(BINARY_CONTENT_TYPE, binary.content_type()));
            }
            if binary.rels().iter().next().is_some() {
                return Err(relerr("ActiveX binary part must not have relationships"));
            }
            if binary.blob().len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            total_binary = total_binary
                .checked_add(binary.blob().len())
                .ok_or_else(|| limit("aggregate ActiveX binary bytes"))?;
            if total_binary > MAX_TOTAL_BINARY {
                return Err(limit("aggregate ActiveX binary bytes"));
            }
            binaries.push(Binary {
                relationship_id: id,
                part_uri: binary_uri,
                bytes: binary.blob().to_vec(),
            });
        }
        binaries.sort_by(|a, b| a.relationship_id.cmp(&b.relationship_id));
        loaded.push(LoadedControl {
            control,
            descriptor_uri,
            descriptor,
            binaries,
            preview,
        });
    }
    Ok(ControlSet { controls: loaded })
}

/// Stores a complete, inert ActiveX graph on a worksheet that has no controls.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
) -> Result<()> {
    let prepared = prepare_graph(package, worksheet_uri, value, true)?;
    install_graph(package, worksheet_uri, prepared)
}

/// Atomically replaces the complete inert ActiveX graph of a worksheet.
///
/// An empty set removes the graph. An in-memory package snapshot is used only
/// for rollback; ActiveX payloads are still copied and never interpreted.
pub fn replace_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
) -> Result<()> {
    if value.controls.is_empty() {
        remove_from_worksheet(package, worksheet_uri)?;
        return Ok(());
    }
    validate_control_set(value)?;
    let snapshot = litchi_opc::PackageWriter::to_bytes(package)?;
    let result = (|| {
        remove_from_worksheet(package, worksheet_uri)?;
        store_on_worksheet(package, worksheet_uri, value)
    })();
    if let Err(error) = result {
        *package = OpcPackage::from_bytes(&snapshot)?;
        return Err(error);
    }
    Ok(())
}

/// Removes the complete ActiveX graph from a worksheet.
///
/// Shared descriptor, binary, or preview parts are retained while any other
/// internal relationship still targets them.
pub fn remove_from_worksheet(package: &mut OpcPackage, worksheet_uri: &PackURI) -> Result<bool> {
    let loaded = load_from_worksheet(package, worksheet_uri)?;
    if loaded.controls.is_empty() {
        return Ok(false);
    }
    let worksheet_xml = package.get_part(worksheet_uri)?.blob().to_vec();
    let updated = replace_controls_xml(&worksheet_xml, &Controls::default())?;
    let control_ids: Vec<String> = loaded
        .controls
        .iter()
        .map(|item| item.control.relationship_id.clone())
        .collect();
    let preview_ids: Vec<String> = loaded
        .controls
        .iter()
        .filter_map(|item| item.preview.as_ref().map(|p| p.relationship_id.clone()))
        .collect();
    let remaining_ids = relationship_ids_in_xml(&updated)?;
    package.unsign();
    {
        let worksheet = package.get_part_mut(worksheet_uri)?;
        for id in &control_ids {
            worksheet.rels_mut().remove(id);
        }
        for id in &preview_ids {
            if !remaining_ids.contains(id) {
                worksheet.rels_mut().remove(id);
            }
        }
        worksheet.set_blob(updated);
    }

    let mut binary_candidates = Vec::new();
    for item in &loaded.controls {
        if !part_has_inbound_relationship(package, &item.descriptor_uri)? {
            package.remove_part(&item.descriptor_uri);
            binary_candidates.extend(item.binaries.iter().map(|b| b.part_uri.clone()));
        }
    }
    for uri in binary_candidates {
        if !part_has_inbound_relationship(package, &uri)? {
            package.remove_part(&uri);
        }
    }
    for preview in loaded
        .controls
        .iter()
        .filter_map(|item| item.preview.as_ref())
    {
        if !part_has_inbound_relationship(package, &preview.part_uri)? {
            package.remove_part(&preview.part_uri);
        }
    }
    Ok(true)
}

struct PreparedGraph {
    worksheet_xml: Vec<u8>,
    strict: bool,
    descriptors: Vec<PreparedDescriptor>,
    resources: Vec<(PackURI, String, Vec<u8>)>,
    worksheet_relationships: Vec<(String, PackURI, bool)>,
}

struct PreparedDescriptor {
    uri: PackURI,
    xml: Vec<u8>,
    relationships: Vec<(String, PackURI)>,
}

fn prepare_graph(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
    require_empty: bool,
) -> Result<PreparedGraph> {
    validate_control_set(value)?;
    let worksheet = package.get_part(worksheet_uri)?;
    if worksheet.content_type() != WORKSHEET_CONTENT_TYPE {
        return Err(content_type(
            WORKSHEET_CONTENT_TYPE,
            worksheet.content_type(),
        ));
    }
    let existing = Controls::parse(worksheet.blob())?;
    if require_empty
        && (!existing.controls.is_empty()
            || worksheet
                .rels()
                .iter()
                .any(|rel| matches!(rel.reltype(), CONTROL_REL | CONTROL_REL_STRICT)))
    {
        return Err(invalid("worksheet already has an ActiveX control graph"));
    }
    let controls = Controls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    };
    let worksheet_xml = replace_controls_xml(worksheet.blob(), &controls)?;
    let strict = controls_span(worksheet.blob())?.strict;

    let mut occupied_ids: HashSet<String> = worksheet
        .rels()
        .iter()
        .map(|r| r.r_id().to_string())
        .collect();
    let mut part_uris = HashSet::new();
    let mut descriptors = Vec::with_capacity(value.controls.len());
    let mut resources = Vec::new();
    let mut worksheet_relationships = Vec::new();
    for item in &value.controls {
        validate_rel_id(&item.control.relationship_id)?;
        if !occupied_ids.insert(item.control.relationship_id.clone()) {
            return Err(relerr("duplicate or occupied worksheet relationship ID"));
        }
        validate_part_location(&item.descriptor_uri, "/xl/activeX/", "ActiveX descriptor")?;
        reserve_new_part(package, &mut part_uris, &item.descriptor_uri)?;
        let descriptor_xml = item.descriptor.to_xml()?;
        let mut expected_ids = HashSet::new();
        collect_binary_ids_descriptor(&item.descriptor, &mut expected_ids)?;
        let actual_ids: HashSet<String> = item
            .binaries
            .iter()
            .map(|binary| binary.relationship_id.clone())
            .collect();
        if expected_ids != actual_ids || actual_ids.len() != item.binaries.len() {
            return Err(relerr(
                "descriptor relationship IDs must exactly match supplied binaries",
            ));
        }
        let mut descriptor_rels = Vec::with_capacity(item.binaries.len());
        for binary in &item.binaries {
            validate_rel_id(&binary.relationship_id)?;
            validate_part_location(&binary.part_uri, "/xl/activeX/", "ActiveX binary")?;
            if binary.bytes.len() > MAX_BINARY {
                return Err(limit("ActiveX binary bytes"));
            }
            reserve_new_part(package, &mut part_uris, &binary.part_uri)?;
            descriptor_rels.push((binary.relationship_id.clone(), binary.part_uri.clone()));
            resources.push((
                binary.part_uri.clone(),
                BINARY_CONTENT_TYPE.into(),
                binary.bytes.clone(),
            ));
        }
        worksheet_relationships.push((
            item.control.relationship_id.clone(),
            item.descriptor_uri.clone(),
            false,
        ));
        match (&item.control.properties, &item.preview) {
            (Some(properties), Some(preview)) => {
                if properties.preview_relationship_id.as_deref()
                    != Some(preview.relationship_id.as_str())
                {
                    return Err(relerr(
                        "control preview relationship ID does not match supplied preview",
                    ));
                }
                validate_rel_id(&preview.relationship_id)?;
                if !occupied_ids.insert(preview.relationship_id.clone()) {
                    return Err(relerr("duplicate or occupied worksheet relationship ID"));
                }
                validate_part_location(&preview.part_uri, "/xl/media/", "ActiveX preview")?;
                if !preview.content_type.starts_with("image/") {
                    return Err(invalid("ActiveX preview content type must be image/*"));
                }
                if preview.bytes.len() > MAX_BINARY {
                    return Err(limit("ActiveX preview image bytes"));
                }
                reserve_new_part(package, &mut part_uris, &preview.part_uri)?;
                worksheet_relationships.push((
                    preview.relationship_id.clone(),
                    preview.part_uri.clone(),
                    true,
                ));
                resources.push((
                    preview.part_uri.clone(),
                    preview.content_type.clone(),
                    preview.bytes.clone(),
                ));
            },
            (Some(properties), None) if properties.preview_relationship_id.is_some() => {
                return Err(relerr("control references a preview that was not supplied"));
            },
            (_, Some(_)) => return Err(relerr("supplied preview is not referenced by controlPr")),
            _ => {},
        }
        descriptors.push(PreparedDescriptor {
            uri: item.descriptor_uri.clone(),
            xml: descriptor_xml,
            relationships: descriptor_rels,
        });
    }
    Ok(PreparedGraph {
        worksheet_xml,
        strict,
        descriptors,
        resources,
        worksheet_relationships,
    })
}

fn install_graph(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    prepared: PreparedGraph,
) -> Result<()> {
    package.unsign();
    for (uri, content_type, bytes) in prepared.resources {
        package.try_add_part(Box::new(BlobPart::new(uri, content_type, bytes)))?;
    }
    for descriptor in prepared.descriptors {
        let mut part = BlobPart::new(
            descriptor.uri.clone(),
            DESCRIPTOR_CONTENT_TYPE.into(),
            descriptor.xml,
        );
        for (id, target) in descriptor.relationships {
            part.rels_mut().try_add_relationship(
                BINARY_REL.into(),
                target.relative_ref(descriptor.uri.base_uri()),
                id,
                TargetMode::Internal,
            )?;
        }
        package.try_add_part(Box::new(part))?;
    }
    let worksheet = package.get_part_mut(worksheet_uri)?;
    for (id, target, preview) in prepared.worksheet_relationships {
        worksheet.rels_mut().try_add_relationship(
            if preview {
                if prepared.strict {
                    IMAGE_REL_STRICT
                } else {
                    IMAGE_REL
                }
            } else if prepared.strict {
                CONTROL_REL_STRICT
            } else {
                CONTROL_REL
            }
            .into(),
            target.relative_ref(worksheet_uri.base_uri()),
            id,
            TargetMode::Internal,
        )?;
    }
    worksheet.set_blob(prepared.worksheet_xml);
    Ok(())
}

fn validate_control_set(value: &ControlSet) -> Result<()> {
    if value.controls.is_empty() || value.controls.len() > MAX_CONTROLS {
        return Err(invalid("ActiveX control set requires 1..65535 controls"));
    }
    validate_controls(&Controls {
        controls: value
            .controls
            .iter()
            .map(|item| item.control.clone())
            .collect(),
    })?;
    let mut total = 0usize;
    for item in &value.controls {
        validate_descriptor(&item.descriptor)?;
        for binary in &item.binaries {
            total = total
                .checked_add(binary.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
        if let Some(preview) = item.preview.as_ref() {
            total = total
                .checked_add(preview.bytes.len())
                .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
        }
    }
    if total > MAX_TOTAL_BINARY {
        return Err(limit("aggregate ActiveX resource bytes"));
    }
    Ok(())
}

fn reserve_new_part(
    package: &OpcPackage,
    reserved: &mut HashSet<PackURI>,
    uri: &PackURI,
) -> Result<()> {
    if reserved.iter().any(|other| other.is_equivalent_to(uri)) {
        return Err(invalid("ActiveX graph contains conflicting part names"));
    }
    package.validate_new_part_name(uri)?;
    reserved.insert(uri.clone());
    Ok(())
}

fn validate_part_location(uri: &PackURI, prefix: &str, kind: &str) -> Result<()> {
    let Some(filename) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{kind} must be stored below {prefix}")));
    };
    if filename.is_empty() || filename.contains('/') {
        return Err(invalid(format!(
            "{kind} must be a direct child of {prefix}"
        )));
    }
    Ok(())
}

fn validate_rel_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    if !matches!(bytes.next(), Some(b'A'..=b'Z' | b'a'..=b'z' | b'_'))
        || !bytes.all(|b| matches!(b, b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'_' | b'.' | b'-'))
    {
        return Err(relerr("relationship ID must be an XML NCName"));
    }
    bounded(value, "relationship ID")
}

fn part_has_inbound_relationship(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::active_x::{
        AX, MAX_CONTROL_NAME_CHARS, MAX_SHAPE_ID, MAX_STRING, REL, REL_STRICT, SML, SML_STRICT,
        X14, XDR_STRICT,
    };
    use litchi_opc::{BlobPart, PackageWriter, Part};

    const WS: &str = "/xl/worksheets/sheet1.xml";
    fn fixture(bytes: &[u8]) -> ControlSet {
        let p = OpcPackage::from_bytes(bytes).unwrap();
        load_from_worksheet(&p, &PackURI::new(WS).unwrap()).unwrap()
    }

    #[test]
    fn libreoffice_stream_init_is_opaque_and_anchored() {
        let set = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/activex_checkbox.xlsx"
        ));
        assert_eq!(set.controls.len(), 1);
        let item = &set.controls[0];
        assert_eq!(item.control.shape_id, 1025);
        assert_eq!(item.descriptor.persistence, Persistence::StreamInit);
        assert_eq!(item.binaries.len(), 1);
        assert_eq!(item.binaries[0].bytes.len(), 116);
        assert_eq!(
            item.control.properties.as_ref().unwrap().anchor.from.column,
            1
        );
        assert_eq!(item.binaries[0].bytes, item.binaries[0].bytes.clone());
    }

    #[test]
    fn libreoffice_radio_buttons_resolve_five_inert_payloads() {
        let set = fixture(include_bytes!(
            "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf111980_radioButtons.xlsx"
        ));
        assert_eq!(set.controls.len(), 5);
        assert!(
            set.controls
                .iter()
                .all(|c| c.descriptor.persistence == Persistence::StreamInit
                    && c.binaries.len() == 1)
        );
    }

    #[test]
    fn poi_property_bag_header_and_footer() {
        for bytes in [
            include_bytes!(
                "../../../../test-data/poi/test-data/spreadsheet/45540_form_Header.xlsx"
            )
            .as_slice(),
            include_bytes!(
                "../../../../test-data/poi/test-data/spreadsheet/45540_form_Footer.xlsx"
            )
            .as_slice(),
        ] {
            let set = fixture(bytes);
            assert_eq!(set.controls.len(), 40);
            assert!(
                set.controls
                    .iter()
                    .all(|c| c.descriptor.persistence == Persistence::PropertyBag
                        && c.binaries.is_empty()
                        && !c.descriptor.properties.is_empty())
            );
        }
    }

    #[test]
    fn strict_nested_mce_controls_roundtrip() {
        let xml = format!(
            r#"<worksheet xmlns="{SML_STRICT}" xmlns:r="{REL_STRICT}" xmlns:xdr="{XDR_STRICT}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{X14}"><mc:AlternateContent><mc:Choice Requires="x14"><controls><mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="7" r:id="rId1" name="safe"><controlPr macro="inert"><anchor moveWithCells="true"><from><xdr:col>1</xdr:col><xdr:colOff>2</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>4</xdr:rowOff></from><to><xdr:col>5</xdr:col><xdr:colOff>6</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>8</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>"#
        );
        let controls = Controls::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            controls.controls[0]
                .properties
                .as_ref()
                .unwrap()
                .macro_name
                .as_deref(),
            Some("inert")
        );
        let canonical = controls.to_xml(true).unwrap();
        assert_eq!(Controls::parse(&canonical).unwrap(), controls);
    }

    #[test]
    fn descriptor_persistence_variants_and_nested_objects() {
        for mode in [
            Persistence::Stream,
            Persistence::StreamInit,
            Persistence::Storage,
        ] {
            let d = Descriptor {
                class_id: "{inert}".into(),
                license: Some("not-used".into()),
                persistence: mode,
                relationship_id: Some("rId1".into()),
                properties: vec![],
            };
            assert_eq!(Descriptor::parse(&d.to_xml().unwrap()).unwrap(), d);
        }
        let d = Descriptor {
            class_id: "not-activated".into(),
            license: None,
            persistence: Persistence::PropertyBag,
            relationship_id: None,
            properties: vec![
                Property {
                    name: "Font".into(),
                    value: None,
                    object: Some(PropertyObject::Font(Font {
                        persistence: Some(Persistence::PropertyBag),
                        relationship_id: None,
                        properties: vec![Property {
                            name: "Name".into(),
                            value: Some("A&B".into()),
                            object: None,
                        }],
                    })),
                },
                Property {
                    name: "Picture".into(),
                    value: None,
                    object: Some(PropertyObject::Picture(Picture {
                        relationship_id: Some("rId2".into()),
                    })),
                },
            ],
        };
        assert_eq!(Descriptor::parse(&d.to_xml().unwrap()).unwrap(), d);
    }

    fn package(external: bool, wrong_type: bool, outbound_binary: bool) -> OpcPackage {
        let worksheet_xml = format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><controls><control shapeId="1" r:id="rId1"/></controls></worksheet>"#).into_bytes();
        let descriptor_xml = format!(r#"<ax:ocx xmlns:ax="{AX}" xmlns:r="{REL}" ax:classid="inert" ax:persistence="persistStreamInit" r:id="rId1"/>"#).into_bytes();
        let mut worksheet = BlobPart::new(
            PackURI::new(WS).unwrap(),
            WORKSHEET_CONTENT_TYPE.into(),
            worksheet_xml,
        );
        worksheet.rels_mut().add_relationship(
            CONTROL_REL.into(),
            if external {
                "https://invalid.example/control".into()
            } else {
                "../activeX/activeX1.xml".into()
            },
            "rId1".into(),
            external,
        );
        let mut descriptor = BlobPart::new(
            PackURI::new("/xl/activeX/activeX1.xml").unwrap(),
            if wrong_type {
                "text/xml".into()
            } else {
                DESCRIPTOR_CONTENT_TYPE.into()
            },
            descriptor_xml,
        );
        descriptor.rels_mut().add_relationship(
            BINARY_REL.into(),
            "activeX1.bin".into(),
            "rId1".into(),
            false,
        );
        let mut binary = BlobPart::new(
            PackURI::new("/xl/activeX/activeX1.bin").unwrap(),
            BINARY_CONTENT_TYPE.into(),
            vec![0, 1, 2, 255],
        );
        if outbound_binary {
            binary.rels_mut().add_relationship(
                IMAGE_REL.into(),
                "../media/x.png".into(),
                "rId9".into(),
                false,
            );
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(worksheet));
        package.add_part(Box::new(descriptor));
        package.add_part(Box::new(binary));
        package
    }

    #[test]
    fn package_validation_and_exact_opaque_roundtrip() {
        let p = package(false, false, false);
        let set = load_from_worksheet(&p, &PackURI::new(WS).unwrap()).unwrap();
        assert_eq!(set.controls[0].binaries[0].bytes, vec![0, 1, 2, 255]);
        assert!(
            load_from_worksheet(&package(true, false, false), &PackURI::new(WS).unwrap()).is_err()
        );
        assert!(
            load_from_worksheet(&package(false, true, false), &PackURI::new(WS).unwrap()).is_err()
        );
        assert!(
            load_from_worksheet(&package(false, false, true), &PackURI::new(WS).unwrap()).is_err()
        );
    }

    #[test]
    fn malformed_and_resource_matrix() {
        assert!(Controls::parse(br#"<!DOCTYPE x><worksheet/>"#).is_err());
        assert!(Controls::parse(format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><controls><control shapeId="0" r:id="x"/></controls></worksheet>"#).as_bytes()).is_err());
        assert!(
            Descriptor::parse(
                format!(r#"<ax:ocx xmlns:ax="{AX}" ax:classid="x" ax:persistence="bad"/>"#)
                    .as_bytes()
            )
            .is_err()
        );
        assert!(Descriptor::parse(format!(r#"<ax:ocx xmlns:ax="{AX}" ax:classid="x" ax:persistence="persistPropertyBag"><ax:ocxPr ax:name="a" ax:value="1"/><ax:ocxPr ax:name="a" ax:value="2"/></ax:ocx>"#).as_bytes()).is_err());
        let huge = "x".repeat(MAX_STRING + 1);
        assert!(
            Descriptor {
                class_id: huge,
                license: None,
                persistence: Persistence::PropertyBag,
                relationship_id: None,
                properties: vec![Property {
                    name: "a".into(),
                    value: None,
                    object: None
                }]
            }
            .to_xml()
            .is_err()
        );
    }

    fn blank_package() -> OpcPackage {
        let xml = format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"/></sheetData><oleObjects/><tableParts count="0"/><extLst/></worksheet>"#
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(WS).unwrap(),
            WORKSHEET_CONTENT_TYPE.into(),
            xml.into_bytes(),
        )));
        package
    }

    fn binary_control(descriptor_uri: &str, binary_uri: &str) -> ControlSet {
        ControlSet {
            controls: vec![LoadedControl {
                control: Control {
                    shape_id: 42,
                    relationship_id: "rIdControl".into(),
                    name: Some("Generated control".into()),
                    properties: Some(ControlProperties {
                        anchor: ObjectAnchor {
                            from: Marker {
                                column: 1,
                                column_offset: 2,
                                row: 3,
                                row_offset: 4,
                            },
                            to: Marker {
                                column: 5,
                                column_offset: 6,
                                row: 7,
                                row_offset: 8,
                            },
                            move_with_cells: Some(true),
                            size_with_cells: Some(false),
                        },
                        locked: Some(true),
                        default_size: None,
                        print: Some(false),
                        disabled: None,
                        recalc_always: None,
                        ui_object: None,
                        auto_fill: None,
                        auto_line: None,
                        auto_picture: None,
                        macro_name: Some("inert_callback_name".into()),
                        alternate_text: Some("generated".into()),
                        preview_relationship_id: Some("rIdPreview".into()),
                    }),
                },
                descriptor_uri: PackURI::new(descriptor_uri).unwrap(),
                descriptor: Descriptor {
                    class_id: "{00000000-0000-0000-0000-000000000000}".into(),
                    license: None,
                    persistence: Persistence::StreamInit,
                    relationship_id: Some("rIdBinary".into()),
                    properties: Vec::new(),
                },
                binaries: vec![Binary {
                    relationship_id: "rIdBinary".into(),
                    part_uri: PackURI::new(binary_uri).unwrap(),
                    bytes: vec![0, 1, 2, 0xff],
                }],
                preview: Some(PreviewImage {
                    relationship_id: "rIdPreview".into(),
                    part_uri: PackURI::new("/xl/media/generated.png").unwrap(),
                    content_type: "image/png".into(),
                    bytes: vec![0x89, b'P', b'N', b'G'],
                }),
            }],
        }
    }

    #[test]
    fn generated_graph_store_reload_and_remove_preserves_unrelated_xml() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let original = package.get_part(&worksheet_uri).unwrap().blob().to_vec();
        let expected = binary_control("/xl/activeX/generated.xml", "/xl/activeX/generated.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &expected).unwrap();
        let xml = std::str::from_utf8(package.get_part(&worksheet_uri).unwrap().blob()).unwrap();
        assert!(xml.contains("<sheetData><row r=\"1\"/></sheetData><oleObjects/><controls"));
        assert!(xml.contains("</controls><tableParts count=\"0\"/><extLst/>"));

        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            load_from_worksheet(&reopened, &worksheet_uri).unwrap(),
            expected
        );

        assert!(remove_from_worksheet(&mut package, &worksheet_uri).unwrap());
        assert_eq!(package.get_part(&worksheet_uri).unwrap().blob(), original);
        assert!(
            package
                .get_part(&PackURI::new("/xl/activeX/generated.xml").unwrap())
                .is_err()
        );
        assert!(!remove_from_worksheet(&mut package, &worksheet_uri).unwrap());
    }

    #[test]
    fn generated_replace_rolls_back_on_conflicting_part_name() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let first = binary_control("/xl/activeX/first.xml", "/xl/activeX/first.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &first).unwrap();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/activeX/occupied.bin").unwrap(),
            "application/octet-stream".into(),
            vec![9],
        )));
        let before_xml = package.get_part(&worksheet_uri).unwrap().blob().to_vec();
        let before_parts = package.part_count();
        let replacement = binary_control("/xl/activeX/second.xml", "/xl/activeX/occupied.bin");
        assert!(replace_on_worksheet(&mut package, &worksheet_uri, &replacement).is_err());
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(package.get_part(&worksheet_uri).unwrap().blob(), before_xml);
        assert_eq!(
            load_from_worksheet(&package, &worksheet_uri).unwrap(),
            first
        );
    }

    #[test]
    fn generated_remove_retains_shared_preview_and_rejects_limits() {
        let mut package = blank_package();
        let worksheet_uri = PackURI::new(WS).unwrap();
        let value = binary_control("/xl/activeX/a.xml", "/xl/activeX/a.bin");
        store_on_worksheet(&mut package, &worksheet_uri, &value).unwrap();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                IMAGE_REL.into(),
                "../media/generated.png".into(),
                "rIdShared".into(),
                TargetMode::Internal,
            )
            .unwrap();
        remove_from_worksheet(&mut package, &worksheet_uri).unwrap();
        assert!(
            package
                .get_part(&PackURI::new("/xl/media/generated.png").unwrap())
                .is_ok()
        );

        let mut invalid = value;
        invalid.controls[0].control.shape_id = MAX_SHAPE_ID + 1;
        assert!(store_on_worksheet(&mut blank_package(), &worksheet_uri, &invalid).is_err());
        invalid.controls[0].control.shape_id = 1;
        invalid.controls[0].control.name = Some("x".repeat(MAX_CONTROL_NAME_CHARS + 1));
        assert!(store_on_worksheet(&mut blank_package(), &worksheet_uri, &invalid).is_err());
    }

    #[test]
    fn direct_xml_mutation_refuses_mce_selected_collection() {
        let xml = format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{X14}"><mc:AlternateContent><mc:Choice Requires="x14"><controls><control shapeId="1" r:id="rId1"/></controls></mc:Choice></mc:AlternateContent></worksheet>"#
        );
        assert!(replace_controls_xml(xml.as_bytes(), &Controls::default()).is_err());
    }
}
