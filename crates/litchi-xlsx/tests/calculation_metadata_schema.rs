use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, TargetMode};
use litchi_xlsx::calculation_properties::{
    Feature, Features, Mode, Properties, ReferenceMode, Snapshot, parse, parse_features,
};
use litchi_xlsx::{Error, Package};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const CALC_FEATURES: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
const CALC_FEATURES_URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";

fn workbook_xml(body: &str) -> Vec<u8> {
    format!(r#"<workbook xmlns="{MAIN}" xmlns:xcalcf="{CALC_FEATURES}">{body}</workbook>"#)
        .into_bytes()
}

fn package_with_workbook_xml(xml: Vec<u8>) -> Package {
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    let workbook = raw
        .main_document_part()
        .expect("main workbook part")
        .partname()
        .clone();
    raw.get_part_mut(&workbook)
        .expect("mutable workbook part")
        .set_blob(xml);
    Package::from_opc(raw).expect("valid workbook fixture")
}

fn reopen(package: &Package) -> Package {
    Package::from_bytes(package.to_bytes().expect("serialize package"))
        .expect("reopen serialized package")
}

fn set_properties(package: &mut Package, properties: Properties) {
    let mut edit = package
        .edit_calculation_metadata()
        .expect("calculation metadata edit");
    assert!(edit.set_properties(properties));
    assert!(edit.commit().expect("metadata commit").changed());
}

fn set_features(package: &mut Package, features: Features) {
    let mut edit = package
        .edit_calculation_metadata()
        .expect("calculation metadata edit");
    assert!(edit.set_features(features));
    assert!(edit.commit().expect("metadata commit").changed());
}

fn properties(snapshot: &Snapshot) -> &Properties {
    snapshot.properties().expect("calcPr properties")
}

fn assert_all_attributes(actual: &Properties, expected_delta: f64) {
    let specified = actual.specified();
    assert_eq!(specified.calculation_id(), Some(u32::MAX));
    assert_eq!(specified.calculation_mode(), Some(Mode::Manual));
    assert_eq!(specified.full_calculation_on_load(), Some(true));
    assert_eq!(specified.reference_mode(), Some(ReferenceMode::R1C1));
    assert_eq!(specified.iterative_calculation(), Some(true));
    assert_eq!(specified.iteration_count(), Some(4_000_000_001));
    assert_eq!(
        specified
            .iteration_delta()
            .expect("authored delta")
            .to_bits(),
        expected_delta.to_bits()
    );
    assert_eq!(specified.full_precision(), Some(false));
    assert_eq!(specified.calculation_completed(), Some(false));
    assert_eq!(specified.calculate_on_save(), Some(false));
    assert_eq!(specified.concurrent_calculation(), Some(false));
    assert_eq!(specified.concurrent_manual_count(), Some(4_000_000_002));
    assert_eq!(specified.force_full_calculation(), Some(true));
}

#[test]
fn all_thirteen_calc_pr_setters_insert_and_round_trip_exact_authored_values() {
    let mut expected = Properties::new();
    expected.set_calculation_id(Some(u32::MAX));
    expected.set_calculation_mode(Some(Mode::Manual));
    expected.set_full_calculation_on_load(Some(true));
    expected.set_reference_mode(Some(ReferenceMode::R1C1));
    expected.set_iterative_calculation(Some(true));
    expected.set_iteration_count(Some(4_000_000_001));
    expected
        .set_iteration_delta(Some(-0.0))
        .expect("xsd:double negative zero");
    expected.set_full_precision(Some(false));
    expected.set_calculation_completed(Some(false));
    expected.set_calculate_on_save(Some(false));
    expected.set_concurrent_calculation(Some(false));
    expected.set_concurrent_manual_count(Some(4_000_000_002));
    expected.set_force_full_calculation(Some(true));

    let mut package = Package::create().expect("fresh package");
    assert!(
        package
            .calculation_metadata()
            .unwrap()
            .properties()
            .is_none()
    );
    set_properties(&mut package, expected.clone());

    let authored = package.calculation_metadata().expect("authored metadata");
    assert_all_attributes(properties(&authored), -0.0);
    assert!(properties(&authored).same_specification(&expected));
    let source = std::str::from_utf8(authored.source_xml()).expect("UTF-8 workbook XML");
    for name in [
        "calcId",
        "calcMode",
        "fullCalcOnLoad",
        "refMode",
        "iterate",
        "iterateCount",
        "iterateDelta",
        "fullPrecision",
        "calcCompleted",
        "calcOnSave",
        "concurrentCalc",
        "concurrentManualCount",
        "forceFullCalc",
    ] {
        assert!(source.contains(&format!(" {name}=\"")), "missing {name}");
    }

    let reopened = reopen(&package);
    let snapshot = reopened
        .calculation_metadata()
        .expect("reopened calculation metadata");
    assert_all_attributes(properties(&snapshot), -0.0);
    assert!(properties(&snapshot).same_specification(&expected));
}

#[test]
fn existing_calc_pr_rewrite_removes_each_authored_attribute_and_then_the_tag() {
    let calc_pr = concat!(
        r#"<calcPr calcId="1" calcMode="autoNoTable" fullCalcOnLoad="1" "#,
        r#"refMode="R1C1" iterate="1" iterateCount="2" iterateDelta="3.5" "#,
        r#"fullPrecision="0" calcCompleted="0" calcOnSave="0" concurrentCalc="0" "#,
        r#"concurrentManualCount="4" forceFullCalc="1"/>"#,
    );
    let mut package = package_with_workbook_xml(workbook_xml(&format!(
        "<sheets/>{calc_pr}<extLst><ext uri=\"{{OPAQUE}}\"><opaque xmlns=\"urn:keep\"/></ext></extLst>"
    )));
    let initial = package.calculation_metadata().expect("initial metadata");
    assert!(initial.properties().is_some());

    let mut edit = package.edit_calculation_metadata().expect("rewrite edit");
    edit.edit_properties(|value| {
        value.set_calculation_id(None);
        value.set_calculation_mode(None);
        value.set_full_calculation_on_load(None);
        value.set_reference_mode(None);
        value.set_iterative_calculation(None);
        value.set_iteration_count(None);
        value.set_iteration_delta(None)?;
        value.set_full_precision(None);
        value.set_calculation_completed(None);
        value.set_calculate_on_save(None);
        value.set_concurrent_calculation(None);
        value.set_concurrent_manual_count(None);
        value.set_force_full_calculation(None);
        Ok(())
    })
    .expect("edit every property");
    assert!(edit.commit().expect("rewrite commit").changed());

    let rewritten = reopen(&package);
    let snapshot = rewritten
        .calculation_metadata()
        .expect("rewritten metadata");
    assert!(properties(&snapshot).same_specification(&Properties::new()));
    let xml = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(xml.contains("<calcPr/>"));
    assert!(xml.contains("<opaque xmlns=\"urn:keep\"/>"));
    for name in [
        "calcId",
        "calcMode",
        "fullCalcOnLoad",
        "refMode",
        "iterate=",
        "iterateCount",
        "iterateDelta",
        "fullPrecision",
        "calcCompleted",
        "calcOnSave",
        "concurrentCalc",
        "concurrentManualCount",
        "forceFullCalc",
    ] {
        assert!(!xml.contains(name), "attribute {name} survived removal");
    }

    let mut rewritten = rewritten;
    let mut remove = rewritten.edit_calculation_metadata().expect("remove edit");
    assert!(remove.remove_properties());
    assert!(remove.commit().expect("remove commit").changed());
    let removed = reopen(&rewritten);
    let snapshot = removed.calculation_metadata().expect("removed metadata");
    assert!(snapshot.properties().is_none());
    let xml = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(!xml.contains("calcPr"));
    assert!(xml.contains("<opaque xmlns=\"urn:keep\"/>"));
}

#[test]
fn xsd_atomic_whitespace_and_leading_zero_unsigned_are_canonicalized_and_plus_rejected() {
    let xml = workbook_xml(concat!(
        r#"<calcPr calcId=" &#x9;00042&#xD; " calcMode="  manual  " "#,
        r#"fullCalcOnLoad=" true " refMode=" R1C1 " iterate=" 1 " "#,
        r#"iterateCount=" 0007 " iterateDelta=" +1.25E+2 " fullPrecision=" false " "#,
        r#"calcCompleted=" 0 " calcOnSave=" true " concurrentCalc=" false " "#,
        r#"concurrentManualCount=" &#xA;0009 " forceFullCalc=" 1 "/>"#,
    ));
    let parsed = parse(&xml).expect("xsd lexical forms").expect("calcPr");
    let specified = parsed.specified();
    assert_eq!(specified.calculation_id(), Some(42));
    assert_eq!(specified.calculation_mode(), Some(Mode::Manual));
    assert_eq!(specified.full_calculation_on_load(), Some(true));
    assert_eq!(specified.reference_mode(), Some(ReferenceMode::R1C1));
    assert_eq!(specified.iterative_calculation(), Some(true));
    assert_eq!(specified.iteration_count(), Some(7));
    assert_eq!(specified.iteration_delta(), Some(125.0));
    assert_eq!(specified.full_precision(), Some(false));
    assert_eq!(specified.calculation_completed(), Some(false));
    assert_eq!(specified.calculate_on_save(), Some(true));
    assert_eq!(specified.concurrent_calculation(), Some(false));
    assert_eq!(specified.concurrent_manual_count(), Some(9));
    assert_eq!(specified.force_full_calculation(), Some(true));

    let mut package = Package::create().expect("fresh package");
    set_properties(&mut package, parsed);
    let reopened = reopen(&package);
    let snapshot = reopened.calculation_metadata().expect("canonical metadata");
    let canonical = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(canonical.contains(r#" calcId="42""#));
    assert!(canonical.contains(r#" iterateCount="7""#));
    assert!(canonical.contains(r#" concurrentManualCount="9""#));
    assert!(!canonical.contains("+42"));
    assert!(!canonical.contains("+0007"));
    assert!(!canonical.contains("+9"));

    for attribute in ["calcId", "iterateCount", "concurrentManualCount"] {
        let invalid = workbook_xml(&format!(r#"<calcPr {attribute}=" &#x9;+1&#xD; "/>"#));
        let error = parse(&invalid).expect_err("leading plus is not an unsignedInt lexical form");
        assert!(
            matches!(error, Error::Invalid(_)),
            "{attribute} returned the wrong error class: {error}"
        );
    }
}

#[test]
fn complete_xsd_double_special_value_space_round_trips_through_public_setter() {
    for (lexical, value) in [
        ("INF", f64::INFINITY),
        ("-INF", f64::NEG_INFINITY),
        ("NaN", f64::NAN),
    ] {
        let parsed = parse(&workbook_xml(&format!(
            r#"<calcPr iterateDelta="{lexical}"/>"#
        )))
        .expect("special xsd:double lexical")
        .expect("calcPr");
        let parsed_value = parsed.specified().iteration_delta().expect("parsed delta");
        if value.is_nan() {
            assert!(parsed_value.is_nan());
        } else {
            assert_eq!(parsed_value, value);
        }

        let mut properties = Properties::new();
        properties
            .set_iteration_delta(Some(value))
            .expect("public setter accepts xsd:double value space");
        let mut package = Package::create().expect("fresh package");
        set_properties(&mut package, properties);
        let reopened = reopen(&package);
        let snapshot = reopened.calculation_metadata().expect("reopened metadata");
        let actual = snapshot
            .properties()
            .unwrap()
            .specified()
            .iteration_delta()
            .unwrap();
        if value.is_nan() {
            assert!(actual.is_nan());
        } else {
            assert_eq!(actual.to_bits(), value.to_bits());
        }
        assert!(
            std::str::from_utf8(snapshot.source_xml())
                .unwrap()
                .contains(&format!(r#"iterateDelta="{lexical}""#))
        );
    }
}

#[test]
fn malformed_calc_features_schema_boundaries_are_rejected() {
    let target = |payload: &str| {
        format!(r#"<extLst><ext uri="{CALC_FEATURES_URI}">{payload}</ext></extLst>"#)
    };
    let cases = [
        (
            "missing feature name",
            target(r#"<xcalcf:calcFeatures><xcalcf:feature/></xcalcf:calcFeatures>"#),
        ),
        (
            "qualified feature name",
            target(concat!(
                r#"<xcalcf:calcFeatures xmlns:q="urn:qualified">"#,
                r#"<xcalcf:feature q:name="x"/></xcalcf:calcFeatures>"#,
            )),
        ),
        (
            "duplicate feature name",
            target(concat!(
                r#"<xcalcf:calcFeatures><xcalcf:feature name="a" name="b"/>"#,
                r#"</xcalcf:calcFeatures>"#,
            )),
        ),
        (
            "non-leaf feature",
            target(concat!(
                r#"<xcalcf:calcFeatures><xcalcf:feature name="a">"#,
                r#"<xcalcf:child/></xcalcf:feature></xcalcf:calcFeatures>"#,
            )),
        ),
        (
            "duplicate target extension",
            format!(
                concat!(
                    r#"<extLst><ext uri="{0}"><xcalcf:calcFeatures>"#,
                    r#"<xcalcf:feature name="a"/></xcalcf:calcFeatures></ext>"#,
                    r#"<ext uri="{0}"><xcalcf:calcFeatures><xcalcf:feature name="b"/>"#,
                    r#"</xcalcf:calcFeatures></ext></extLst>"#,
                ),
                CALC_FEATURES_URI
            ),
        ),
        (
            "duplicate calcFeatures payload",
            target(concat!(
                r#"<xcalcf:calcFeatures><xcalcf:feature name="a"/></xcalcf:calcFeatures>"#,
                r#"<xcalcf:calcFeatures><xcalcf:feature name="b"/></xcalcf:calcFeatures>"#,
            )),
        ),
        (
            "mixed opaque target payload",
            target(concat!(
                r#"<opaque xmlns="urn:opaque"/>"#,
                r#"<xcalcf:calcFeatures><xcalcf:feature name="a"/></xcalcf:calcFeatures>"#,
            )),
        ),
    ];

    for (label, body) in cases {
        let error = parse_features(&workbook_xml(&body)).expect_err(label);
        if label == "duplicate feature name" {
            assert!(
                matches!(error, Error::Xml(_)),
                "{label} returned the wrong error class: {error}"
            );
        } else {
            assert!(
                matches!(error, Error::Invalid(_)),
                "{label} returned the wrong error class: {error}"
            );
        }
    }
}

#[test]
fn xsd_string_feature_names_preserve_empty_whitespace_and_supplementary_scalars() {
    let names = [
        "",
        "\t",
        "\n",
        "\r",
        "\t\n\r",
        "supplementary-\u{10000}-\u{10ffff}",
    ];
    let features = Features::try_from_vec(
        names
            .iter()
            .map(|name| Feature::new(*name).expect("valid XML 1.0 xsd:string"))
            .collect(),
    )
    .expect("nonempty feature collection");
    let mut package = Package::create().expect("fresh package");
    set_features(&mut package, features);

    let reopened = reopen(&package);
    let snapshot = reopened.calculation_metadata().expect("reopened metadata");
    let actual: Vec<_> = snapshot
        .features()
        .expect("calcFeatures")
        .iter()
        .map(Feature::as_str)
        .collect();
    assert_eq!(actual, names);
    let xml = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(xml.contains("name=\"\""));
    assert!(xml.contains("&#x9;"));
    assert!(xml.contains("&#xA;"));
    assert!(xml.contains("&#xD;"));
    assert!(xml.contains("\u{10000}"));
    assert!(xml.contains("\u{10ffff}"));

    let parsed = parse_features(&workbook_xml(&format!(
        concat!(
            r#"<extLst><ext uri="{}"><xcalcf:calcFeatures>"#,
            r#"<xcalcf:feature name="&#x9;&#xA;&#xD;"/>"#,
            r#"</xcalcf:calcFeatures></ext></extLst>"#,
        ),
        CALC_FEATURES_URI
    )))
    .expect("parse xsd:string whitespace")
    .expect("calcFeatures");
    assert_eq!(parsed.get(0).unwrap().as_str(), "\t\n\r");
}

#[test]
fn strict_prefixed_noncanonical_workbook_writes_and_reopens_calc_pr() {
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    let workbook_name = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let strict_xml = std::str::from_utf8(raw.get_part(&workbook_name).unwrap().blob())
        .unwrap()
        .replace(MAIN, STRICT_MAIN)
        .replace(RELS, STRICT_RELS)
        .replacen(
            "<workbook ",
            &format!(r#"<s:workbook xmlns:s="{STRICT_MAIN}" "#),
            1,
        )
        .replace("</workbook>", "</s:workbook>")
        .into_bytes();
    let main = raw.get_part_mut(&workbook_name).expect("mutable workbook");
    main.set_blob(strict_xml);
    main.rels_mut()
        .remove("rId1")
        .expect("worksheet relationship");
    main.rels_mut()
        .try_add_relationship(
            rt::STRICT_WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("strict worksheet relationship");
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
    let strict_sheet = std::str::from_utf8(raw.get_part(&worksheet).unwrap().blob())
        .unwrap()
        .replace(MAIN, STRICT_MAIN)
        .into_bytes();
    raw.get_part_mut(&worksheet).unwrap().set_blob(strict_sheet);

    let mut package = Package::from_opc(raw).expect("strict package");
    set_properties(
        &mut package,
        Properties::new()
            .with_calculation_id(Some(77))
            .with_reference_mode(Some(ReferenceMode::R1C1)),
    );
    let reopened = reopen(&package);
    let snapshot = reopened.calculation_metadata().expect("strict metadata");
    assert_eq!(properties(&snapshot).specified().calculation_id(), Some(77));
    assert_eq!(
        properties(&snapshot).specified().reference_mode(),
        Some(ReferenceMode::R1C1)
    );
    let xml = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(xml.contains(&format!(r#"<s:workbook xmlns:s="{STRICT_MAIN}""#)));
    assert!(xml.contains("<s:calcPr"));
    assert!(!xml.contains(MAIN));
}

#[test]
fn strict_prefixed_workbook_inserts_calc_features_with_strict_core_qnames() {
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    let workbook_name = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let strict_xml = std::str::from_utf8(raw.get_part(&workbook_name).unwrap().blob())
        .unwrap()
        .replace(MAIN, STRICT_MAIN)
        .replace(RELS, STRICT_RELS)
        .replacen(
            "<workbook ",
            &format!(r#"<s:workbook xmlns:s="{STRICT_MAIN}" "#),
            1,
        )
        .replace("</workbook>", "</s:workbook>")
        .into_bytes();
    let main = raw.get_part_mut(&workbook_name).expect("mutable workbook");
    main.set_blob(strict_xml);
    main.rels_mut()
        .remove("rId1")
        .expect("worksheet relationship");
    main.rels_mut()
        .try_add_relationship(
            rt::STRICT_WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rId1".to_owned(),
            TargetMode::Internal,
        )
        .expect("strict worksheet relationship");
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
    let strict_sheet = std::str::from_utf8(raw.get_part(&worksheet).unwrap().blob())
        .unwrap()
        .replace(MAIN, STRICT_MAIN)
        .into_bytes();
    raw.get_part_mut(&worksheet).unwrap().set_blob(strict_sheet);

    let mut package = Package::from_opc(raw).expect("strict package");
    let absent = package.calculation_metadata().expect("absent metadata");
    assert!(absent.properties().is_none());
    assert!(absent.features().is_none());
    let names = ["strict<&\"'>", "", "supplementary-\u{10000}"];
    let expected = Features::try_from_vec(
        names
            .iter()
            .map(|name| Feature::new(*name).expect("inert feature name"))
            .collect(),
    )
    .expect("nonempty features");
    set_features(&mut package, expected.clone());

    let reopened = reopen(&package);
    let snapshot = reopened.calculation_metadata().expect("strict metadata");
    assert!(snapshot.properties().is_none());
    assert_eq!(snapshot.features(), Some(&expected));
    let xml = std::str::from_utf8(snapshot.source_xml()).expect("UTF-8 workbook XML");
    assert!(xml.contains(&format!(r#"<s:workbook xmlns:s="{STRICT_MAIN}""#)));
    assert!(xml.contains("<s:extLst>"));
    assert!(xml.contains(&format!(r#"<s:ext uri="{CALC_FEATURES_URI}">"#)));
    assert!(xml.contains(&format!(
        r#"<xcalcf:calcFeatures xmlns:xcalcf="{CALC_FEATURES}">"#
    )));
    assert!(xml.contains(r#"<xcalcf:feature name="strict&lt;&amp;&quot;"#));
    assert!(xml.contains(r#"<xcalcf:feature name=""/>"#));
    assert!(xml.contains("supplementary-\u{10000}"));
    assert!(xml.contains("</s:ext></s:extLst></s:workbook>"));
    assert!(!xml.contains(MAIN));
}

#[test]
fn mce_selected_calc_features_are_readable_but_changed_edit_is_refused() {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let mut raw = Package::create().expect("minimal package").into_plain_opc();
    let workbook_name = raw.main_document_part().unwrap().partname().clone();
    let xml = std::str::from_utf8(raw.get_part(&workbook_name).unwrap().blob())
        .unwrap()
        .replacen("<workbook ", &format!(r#"<workbook xmlns:mc="{MC}" "#), 1)
        .replace(
            "</workbook>",
            &format!(
                concat!(
                    r#"<mc:AlternateContent><mc:Choice Requires="future" xmlns:future="urn:future">"#,
                    r#"<extLst><ext uri="{0}"><future:opaque/></ext></extLst></mc:Choice>"#,
                    r#"<mc:Fallback><extLst><ext uri="{0}">"#,
                    r#"<xcalcf:calcFeatures xmlns:xcalcf="{1}"><xcalcf:feature name="selected"/>"#,
                    r#"</xcalcf:calcFeatures></ext></extLst></mc:Fallback></mc:AlternateContent>"#,
                    r#"</workbook>"#,
                ),
                CALC_FEATURES_URI, CALC_FEATURES
            ),
        )
        .into_bytes();
    raw.get_part_mut(&workbook_name)
        .expect("mutable workbook")
        .set_blob(xml.clone());
    let mut package = Package::from_opc(raw).expect("MCE fixture");
    let snapshot = package.calculation_metadata().expect("projected metadata");
    assert_eq!(
        snapshot.features().unwrap().get(0).unwrap().as_str(),
        "selected"
    );

    let mut edit = package.edit_calculation_metadata().expect("feature edit");
    assert!(edit.set_features(Features::new(Feature::new("changed").unwrap())));
    let error = edit.commit().expect_err("projected edit must be refused");
    assert!(matches!(error, Error::Invalid(_)));
    assert!(error.to_string().contains("projected through MCE"));
    assert_eq!(package.calculation_metadata().unwrap().source_xml(), xml);
}
