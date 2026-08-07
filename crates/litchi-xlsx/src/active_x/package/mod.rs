//! OPC package graph lifecycle for inert ActiveX resources.

mod load;
mod transaction;

use super::codec::{relationship_ids_in_xml, replace_controls_xml};
use super::model::*;
use super::validation::validate_control_set;
use super::{
    CONTROL_REL, CONTROL_REL_STRICT, Result, WORKSHEET_CONTENT_TYPE, content_type, relerr,
};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::HashSet;

#[cfg(test)]
use super::{BINARY_CONTENT_TYPE, BINARY_REL, DESCRIPTOR_CONTENT_TYPE, IMAGE_REL};
#[cfg(test)]
use litchi_opc::TargetMode;

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
        loaded.push(load::load_control(
            package,
            worksheet,
            control,
            &mut total_binary,
        )?);
    }
    Ok(ControlSet { controls: loaded })
}

/// Stores a complete, inert ActiveX graph on a worksheet that has no controls.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_uri: &PackURI,
    value: &ControlSet,
) -> Result<()> {
    transaction::store_on_worksheet(package, worksheet_uri, value)
}

/// Atomically replaces the complete inert ActiveX graph of a worksheet.
///
/// An empty set removes the graph. A typed package clone is used only for
/// rollback; ActiveX payloads are still copied and never interpreted.
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
    let snapshot = package.clone();
    let result = (|| {
        remove_from_worksheet(package, worksheet_uri)?;
        store_on_worksheet(package, worksheet_uri, value)
    })();
    if let Err(error) = result {
        *package = snapshot;
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
        if !transaction::part_has_inbound_relationship(package, &item.descriptor_uri)? {
            package.remove_part(&item.descriptor_uri);
            binary_candidates.extend(item.binaries.iter().map(|b| b.part_uri.clone()));
        }
    }
    for uri in binary_candidates {
        if !transaction::part_has_inbound_relationship(package, &uri)? {
            package.remove_part(&uri);
        }
    }
    for preview in loaded
        .controls
        .iter()
        .filter_map(|item| item.preview.as_ref())
    {
        if !transaction::part_has_inbound_relationship(package, &preview.part_uri)? {
            package.remove_part(&preview.part_uri);
        }
    }
    Ok(true)
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
            "../../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/activex_checkbox.xlsx"
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
            "../../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf111980_radioButtons.xlsx"
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
                "../../../../../test-data/poi/test-data/spreadsheet/45540_form_Header.xlsx"
            )
            .as_slice(),
            include_bytes!(
                "../../../../../test-data/poi/test-data/spreadsheet/45540_form_Footer.xlsx"
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
