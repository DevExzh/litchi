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

#[test]
fn rich_part_content_types_fold_ascii_case_without_widening_the_profile() {
    for kind in [
        Kind::Data,
        Kind::Structures,
        Kind::Arrays,
        Kind::Styles,
        Kind::SupportingData,
        Kind::SupportingStructures,
        Kind::Types,
        Kind::WebImages,
        Kind::Relationships,
        Kind::FeatureBags,
    ] {
        for value in [
            kind.content_type().to_owned(),
            kind.content_type().to_ascii_lowercase(),
            kind.content_type().to_ascii_uppercase(),
        ] {
            assert_eq!(Kind::from_content_type(&value), Some(kind));
        }
        for invalid in [
            format!(" {}", kind.content_type()),
            format!("{};charset=utf-8", kind.content_type()),
            format!("{}x", kind.content_type()),
        ] {
            assert_eq!(Kind::from_content_type(&invalid), None);
        }
    }
}

#[test]
fn rich_relationship_and_style_fragment_writers_close_their_start_tags() {
    use super::codec::{
        parse_dxf_complement, parse_rich_value_rels, parse_xf_complement, write_dxf_complement,
        write_rich_value_rels, write_xf_complement,
    };
    let rels = super::RichValueRels {
        ids: vec!["rId&quoted\"".into()],
        extension_list: None,
        opaque: Vec::new(),
    };
    assert_eq!(
        parse_rich_value_rels(&write_rich_value_rels(&rels).unwrap()).unwrap(),
        rels
    );
    let xf=parse_xf_complement(format!(r#"<f:xfComplement xmlns:f="{FEATURE_BAG}" i="2"><v:keep xmlns:v="urn:vendor">payload</v:keep></f:xfComplement>"#).as_bytes()).unwrap();
    let dxf=parse_dxf_complement(format!(r#"<f:DXFComplement xmlns:f="{FEATURE_BAG}" i="3"><v:keep xmlns:v="urn:vendor">payload</v:keep></f:DXFComplement>"#).as_bytes()).unwrap();
    assert_eq!(xf.index, 2);
    assert_eq!(dxf.index, 3);
    assert_eq!(
        parse_xf_complement(&write_xf_complement(&xf).unwrap()).unwrap(),
        xf
    );
    assert_eq!(
        parse_dxf_complement(&write_dxf_complement(&dxf).unwrap()).unwrap(),
        dxf
    );
}

/// Every sequence of at most `max_len` child indices below `count`,
/// shortest first.
fn sequences(count: usize, max_len: usize) -> Vec<Vec<usize>> {
    let mut all = vec![Vec::new()];
    let mut start = 0;
    for _ in 0..max_len {
        let end = all.len();
        for index in start..end {
            for child in 0..count {
                let mut next = all[index].clone();
                next.push(child);
                all.push(next);
            }
        }
        start = end;
    }
    all
}

fn assert_rows(actual: &[(String, String)], expected: &[(&str, &str)]) {
    let same = actual.len() == expected.len()
        && actual.iter().zip(expected).all(
            |((label, outcome), (expected_label, expected_outcome))| {
                label == expected_label && outcome == expected_outcome
            },
        );
    let observed: String = actual
        .iter()
        .map(|(label, outcome)| format!("({label:?}, {outcome:?}),\n"))
        .collect();
    assert!(same, "observed rows:\n{observed}");
}

#[test]
fn feature_property_bag_child_order_matrix_is_stable() {
    let children = [
        ("bagExt", "<fpb:bagExt/>"),
        ("bag", "<fpb:bag type=\"Checkbox\"/>"),
        ("extLst", "<x:extLst/>"),
    ];
    let marked = (
        "marked extLst",
        "<x:extLst><x:ext uri=\"urn:second\"/></x:extLst>",
    );
    let mut cases: Vec<Vec<(&str, &str)>> = sequences(children.len(), 3)
        .into_iter()
        .map(|sequence| sequence.into_iter().map(|index| children[index]).collect())
        .collect();
    cases.push(vec![("unknown", "<fpb:unknown/>")]);
    cases.push(vec![
        ("unknown", "<fpb:unknown/>"),
        children[0],
        children[1],
        children[2],
    ]);
    cases.push(vec![("core bag", "<x:bag/>")]);
    cases.push(vec![children[2], marked]);
    cases.push(vec![marked, children[2]]);
    let actual: Vec<(String, String)> = cases
        .iter()
        .map(|case| {
            let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
            let body: String = case.iter().map(|(_, xml)| *xml).collect();
            let xml = format!(
                "<fpb:FeaturePropertyBags xmlns:fpb=\"{FEATURE_BAG}\" xmlns:x=\"{SPREADSHEETML}\">{body}</fpb:FeaturePropertyBags>"
            );
            let outcome = match parse_feature_property_bags(xml.as_bytes()) {
                Ok(bags) => {
                    let extension_list = match &bags.extension_list {
                        None => "none",
                        Some(list) if String::from_utf8_lossy(&list.xml).contains("urn:second") => {
                            "marked"
                        },
                        Some(_) => "plain",
                    };
                    format!(
                        "ok bagExt={} bag={} extLst={extension_list} opaque={}",
                        bags.bag_extensions.len(),
                        bags.values.len(),
                        bags.opaque.len()
                    )
                },
                Err(error) => format!("{error:?}"),
            };
            (label.join(","), outcome)
        })
        .collect();
    assert_rows(
        &actual,
        &[
            ("", "ok bagExt=0 bag=0 extLst=none opaque=0"),
            ("bagExt", "ok bagExt=1 bag=0 extLst=none opaque=0"),
            ("bag", "ok bagExt=0 bag=1 extLst=none opaque=0"),
            ("extLst", "ok bagExt=0 bag=0 extLst=plain opaque=0"),
            ("bagExt,bagExt", "ok bagExt=2 bag=0 extLst=none opaque=0"),
            ("bagExt,bag", "ok bagExt=1 bag=1 extLst=none opaque=0"),
            ("bagExt,extLst", "ok bagExt=1 bag=0 extLst=plain opaque=0"),
            ("bag,bagExt", "ok bagExt=0 bag=1 extLst=none opaque=1"),
            ("bag,bag", "ok bagExt=0 bag=2 extLst=none opaque=0"),
            ("bag,extLst", "ok bagExt=0 bag=1 extLst=plain opaque=0"),
            ("extLst,bagExt", "ok bagExt=0 bag=0 extLst=plain opaque=1"),
            ("extLst,bag", "ok bagExt=0 bag=0 extLst=plain opaque=1"),
            ("extLst,extLst", "ok bagExt=0 bag=0 extLst=plain opaque=0"),
            (
                "bagExt,bagExt,bagExt",
                "ok bagExt=3 bag=0 extLst=none opaque=0",
            ),
            (
                "bagExt,bagExt,bag",
                "ok bagExt=2 bag=1 extLst=none opaque=0",
            ),
            (
                "bagExt,bagExt,extLst",
                "ok bagExt=2 bag=0 extLst=plain opaque=0",
            ),
            (
                "bagExt,bag,bagExt",
                "ok bagExt=1 bag=1 extLst=none opaque=1",
            ),
            ("bagExt,bag,bag", "ok bagExt=1 bag=2 extLst=none opaque=0"),
            (
                "bagExt,bag,extLst",
                "ok bagExt=1 bag=1 extLst=plain opaque=0",
            ),
            (
                "bagExt,extLst,bagExt",
                "ok bagExt=1 bag=0 extLst=plain opaque=1",
            ),
            (
                "bagExt,extLst,bag",
                "ok bagExt=1 bag=0 extLst=plain opaque=1",
            ),
            (
                "bagExt,extLst,extLst",
                "ok bagExt=1 bag=0 extLst=plain opaque=0",
            ),
            (
                "bag,bagExt,bagExt",
                "ok bagExt=0 bag=1 extLst=none opaque=2",
            ),
            ("bag,bagExt,bag", "ok bagExt=0 bag=2 extLst=none opaque=1"),
            (
                "bag,bagExt,extLst",
                "ok bagExt=0 bag=1 extLst=plain opaque=1",
            ),
            ("bag,bag,bagExt", "ok bagExt=0 bag=2 extLst=none opaque=1"),
            ("bag,bag,bag", "ok bagExt=0 bag=3 extLst=none opaque=0"),
            ("bag,bag,extLst", "ok bagExt=0 bag=2 extLst=plain opaque=0"),
            (
                "bag,extLst,bagExt",
                "ok bagExt=0 bag=1 extLst=plain opaque=1",
            ),
            ("bag,extLst,bag", "ok bagExt=0 bag=1 extLst=plain opaque=1"),
            (
                "bag,extLst,extLst",
                "ok bagExt=0 bag=1 extLst=plain opaque=0",
            ),
            (
                "extLst,bagExt,bagExt",
                "ok bagExt=0 bag=0 extLst=plain opaque=2",
            ),
            (
                "extLst,bagExt,bag",
                "ok bagExt=0 bag=0 extLst=plain opaque=2",
            ),
            (
                "extLst,bagExt,extLst",
                "ok bagExt=0 bag=0 extLst=plain opaque=1",
            ),
            (
                "extLst,bag,bagExt",
                "ok bagExt=0 bag=0 extLst=plain opaque=2",
            ),
            ("extLst,bag,bag", "ok bagExt=0 bag=0 extLst=plain opaque=2"),
            (
                "extLst,bag,extLst",
                "ok bagExt=0 bag=0 extLst=plain opaque=1",
            ),
            (
                "extLst,extLst,bagExt",
                "ok bagExt=0 bag=0 extLst=plain opaque=1",
            ),
            (
                "extLst,extLst,bag",
                "ok bagExt=0 bag=0 extLst=plain opaque=1",
            ),
            (
                "extLst,extLst,extLst",
                "ok bagExt=0 bag=0 extLst=plain opaque=0",
            ),
            ("unknown", "ok bagExt=0 bag=0 extLst=none opaque=1"),
            (
                "unknown,bagExt,bag,extLst",
                "ok bagExt=1 bag=1 extLst=plain opaque=1",
            ),
            ("core bag", "ok bagExt=0 bag=0 extLst=none opaque=1"),
            (
                "extLst,marked extLst",
                "ok bagExt=0 bag=0 extLst=marked opaque=0",
            ),
            (
                "marked extLst,extLst",
                "ok bagExt=0 bag=0 extLst=plain opaque=0",
            ),
        ],
    );
}
