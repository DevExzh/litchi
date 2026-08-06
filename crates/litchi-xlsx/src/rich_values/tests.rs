use super::codec::{
    parse_arrays, parse_data, parse_feature_property_bags, parse_structures, write_arrays,
    write_feature_property_bags,
};
use super::model::{CheckboxState, Mode};
use super::package::{
    Document, FEATURE_PROPERTY_BAG_CONTENT_TYPE, Kind, RICH_VALUE_DATA_CONTENT_TYPE,
    RICH_VALUE_RELATIONSHIPS_CONTENT_TYPE, RICH_VALUE_STRUCTURE_CONTENT_TYPE, load,
};
use super::{FEATURE_BAG, RICH_DATA, RICH_DATA_2, RICH_VALUE_REL, SPREADSHEETML};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part as _};

#[test]
fn feature_property_bags_keep_checkbox_chain_and_unknown_xml() {
    let xml = format!(
        r#"<fpb:FeaturePropertyBags xmlns:fpb="{FEATURE_BAG}" xmlns:x="{SPREADSHEETML}" count="4">
            <fpb:bag type="Checkbox">
                <fpb:i k="default">1</fpb:i>
                <v:future xmlns:v="urn:vendor:future"><v:value>kept</v:value></v:future>
            </fpb:bag>
            <fpb:bag type="XFControls"><fpb:bagId k="CellControl">0</fpb:bagId></fpb:bag>
            <fpb:bag type="XFComplement"><fpb:bagId k="XFControls">1</fpb:bagId></fpb:bag>
            <fpb:bag type="XFComplements">
                <fpb:a k="MappedFeaturePropertyBags"><fpb:bagId>2</fpb:bagId></fpb:a>
            </fpb:bag>
            <x:extLst><x:ext uri="urn:future"/></x:extLst>
        </fpb:FeaturePropertyBags>"#
    );

    let bags = parse_feature_property_bags(xml.as_bytes()).unwrap();
    assert_eq!(
        bags.checkbox(0).unwrap().unwrap().default,
        CheckboxState::Checked
    );
    assert!(
        bags.get(3)
            .unwrap()
            .property("MappedFeaturePropertyBags")
            .is_some()
    );

    let rewritten = write_feature_property_bags(&bags).unwrap();
    assert!(String::from_utf8_lossy(&rewritten).contains("future"));
    let round_trip = parse_feature_property_bags(&rewritten).unwrap();
    assert_eq!(round_trip, bags);
}

#[test]
fn rich_data_structures_and_arrays_are_typed_and_bounded() {
    let structures_xml = format!(
        r#"<rd:rvStructures xmlns:rd="{RICH_DATA}" xmlns:x="{SPREADSHEETML}" count="1">
            <rd:s t="entity"><rd:k n="name" t="s"/></rd:s>
            <x:extLst><x:ext uri="urn:structure-future"/></x:extLst>
        </rd:rvStructures>"#
    );
    let data_xml = format!(
        r#"<rd:rvData xmlns:rd="{RICH_DATA}" xmlns:x="{SPREADSHEETML}" count="1">
            <rd:rv s="0"><rd:fb t="s">Ada</rd:fb><rd:v>Ada</rd:v></rd:rv>
        </rd:rvData>"#
    );
    let arrays_xml = format!(
        r#"<rd2:arrayData xmlns:rd2="{RICH_DATA_2}" xmlns:x="{SPREADSHEETML}" count="1">
            <rd2:a r="1" c="2"><rd2:v t="s">Ada</rd2:v><rd2:v t="i">7</rd2:v></rd2:a>
        </rd2:arrayData>"#
    );

    let structures = parse_structures(structures_xml.as_bytes()).unwrap();
    let data = parse_data(data_xml.as_bytes()).unwrap();
    let arrays = parse_arrays(arrays_xml.as_bytes()).unwrap();
    assert_eq!(structures.values[0].keys[0].name, "name");
    assert_eq!(data.values[0].values, vec!["Ada"]);
    assert_eq!(arrays.values[0].values.len(), 2);

    assert_eq!(
        parse_structures(&super::codec::write_structures(&structures).unwrap()).unwrap(),
        structures
    );
    assert_eq!(
        parse_arrays(&write_arrays(&arrays).unwrap()).unwrap(),
        arrays
    );
}

#[test]
fn package_snapshot_preserves_typed_parts_and_relationship_topology() {
    let data_xml = format!(
        r#"<rd:rvData xmlns:rd="{RICH_DATA}" count="1"><rd:rv s="0"><rd:v>Ada</rd:v></rd:rv></rd:rvData>"#
    );
    let structures_xml = format!(
        r#"<rd:rvStructures xmlns:rd="{RICH_DATA}" count="1"><rd:s t="entity"><rd:k n="name" t="s"/></rd:s></rd:rvStructures>"#
    );
    let relationships_xml = format!(
        r#"<rvr:richValueRels xmlns:rvr="{RICH_VALUE_REL}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rvr:rel r:id="rIdImage"/></rvr:richValueRels>"#
    );
    let feature_bags_xml =
        format!(r#"<fpb:FeaturePropertyBags xmlns:fpb="{FEATURE_BAG}" count="0"/>"#);

    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/richData.xml").unwrap(),
            RICH_VALUE_DATA_CONTENT_TYPE.into(),
            data_xml.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/richStructures.xml").unwrap(),
            RICH_VALUE_STRUCTURE_CONTENT_TYPE.into(),
            structures_xml.into_bytes(),
        )))
        .unwrap();
    let mut rich_relationships = BlobPart::new(
        PackURI::new("/xl/richValueRels.xml").unwrap(),
        RICH_VALUE_RELATIONSHIPS_CONTENT_TYPE.into(),
        relationships_xml.into_bytes(),
    );
    rich_relationships
        .rels_mut()
        .try_add_relationship(
            "urn:test:image".into(),
            "media/image1.png".into(),
            "rIdImage".into(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(rich_relationships)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/featurePropertyBags.xml").unwrap(),
            FEATURE_PROPERTY_BAG_CONTENT_TYPE.into(),
            feature_bags_xml.into_bytes(),
        )))
        .unwrap();
    let mut worksheet = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
        "application/xml".into(),
        b"<worksheet/>".to_vec(),
    );
    worksheet
        .rels_mut()
        .try_add_relationship(
            "urn:test:rich-data".into(),
            "../richData.xml".into(),
            "rIdRich".into(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(worksheet)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/media/image1.png").unwrap(),
            "image/png".into(),
            vec![0x89, b'P', b'N', b'G'],
        )))
        .unwrap();

    let snapshot = load(&package).unwrap();
    assert_eq!(snapshot.parts().len(), 4);
    assert!(matches!(
        snapshot.part(Kind::FeatureBags).unwrap().document(),
        Document::FeatureBags(_)
    ));
    let rich_relationships = snapshot.part(Kind::Relationships).unwrap();
    assert_eq!(
        rich_relationships.relationships()[0]
            .resolved_target
            .as_deref(),
        Some("/xl/media/image1.png")
    );
    assert_eq!(rich_relationships.relationships()[0].mode, Mode::Internal);
    assert!(snapshot.topology().iter().any(|link| {
        link.source.as_deref() == Some("/xl/worksheets/sheet1.xml")
            && link.id == "rIdRich"
            && link.resolved_target.as_deref() == Some("/xl/richData.xml")
    }));
}
