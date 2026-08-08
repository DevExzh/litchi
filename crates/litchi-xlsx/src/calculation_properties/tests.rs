use super::codec::MAX_XML_BYTES;
use std::borrow::Cow;

use super::{
    Feature, Features, Limits, Mode, Properties, ReferenceMode, inspect, invalidate_formulas,
    parse, parse_features, parse_features_with_limits, rewrite,
};
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn parse_xml(child: &str) -> Result<Option<Properties>> {
    parse(format!(r#"<workbook xmlns="{NS}">{child}</workbook>"#).as_bytes())
}

#[test]
fn parses_all_attributes_and_effective_defaults() {
    let value = parse_xml(concat!(
        r#"<calcPr calcId="42" calcMode="autoNoTable" fullCalcOnLoad="1" "#,
        r#"refMode="R1C1" iterate="true" iterateCount="250" iterateDelta="1E-4" "#,
        r#"fullPrecision="0" calcCompleted="false" calcOnSave="0" concurrentCalc="false" "#,
        r#"concurrentManualCount="6" forceFullCalc="true"/>"#,
    ))
    .unwrap()
    .unwrap();
    assert_eq!(value.calculation_id(), 42);
    assert_eq!(value.calculation_mode(), Mode::AutomaticExceptTables);
    assert!(value.full_calculation_on_load());
    assert_eq!(value.reference_mode(), ReferenceMode::R1C1);
    assert!(value.iterative_calculation());
    assert_eq!(value.iteration_count(), 250);
    assert_eq!(value.iteration_delta(), 0.0001);
    assert!(!value.full_precision());
    assert!(!value.calculation_completed());
    assert!(!value.calculate_on_save());
    assert!(!value.concurrent_calculation());
    assert_eq!(value.concurrent_manual_count(), Some(6));
    assert!(value.force_full_calculation());

    let defaults = parse_xml("<calcPr/>").unwrap().unwrap();
    assert_eq!(defaults.calculation_id(), 0);
    assert_eq!(defaults.calculation_mode(), Mode::Automatic);
    assert!(!defaults.full_calculation_on_load());
    assert_eq!(defaults.reference_mode(), ReferenceMode::A1);
    assert!(!defaults.iterative_calculation());
    assert_eq!(defaults.iteration_count(), 100);
    assert_eq!(defaults.iteration_delta(), 0.001);
    assert!(defaults.full_precision());
    assert!(defaults.calculation_completed());
    assert!(defaults.calculate_on_save());
    assert!(defaults.concurrent_calculation());
    assert_eq!(defaults.concurrent_manual_count(), None);
    assert!(!defaults.force_full_calculation());
}

#[test]
fn supports_strict_namespace_and_mce_fallback() {
    let strict = br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><calcPr calcMode="manual"/></workbook>"#;
    assert_eq!(
        parse(strict).unwrap().unwrap().calculation_mode(),
        Mode::Manual
    );
    let mce = format!(
        concat!(
            r#"<workbook xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:calcPr/></mc:Choice><mc:Fallback>"#,
            r#"<calcPr refMode="R1C1"/></mc:Fallback></mc:AlternateContent></workbook>"#,
        ),
        NS
    );
    assert_eq!(
        parse(mce.as_bytes()).unwrap().unwrap().reference_mode(),
        ReferenceMode::R1C1
    );
}

#[test]
fn rejects_invalid_values_structure_and_attributes() {
    for child in [
        r#"<calcPr calcMode="sometimes"/>"#,
        r#"<calcPr refMode="A2"/>"#,
        r#"<calcPr iterate="yes"/>"#,
        r#"<calcPr iterateCount="-1"/>"#,
        r#"<calcPr mystery="1"/>"#,
        r#"<calcPr><child/></calcPr>"#,
        r#"<wrapper><calcPr calcId="1"/></wrapper>"#,
    ] {
        let result = parse_xml(child);
        if child.starts_with("<wrapper>") {
            assert!(result.unwrap().is_none());
        } else {
            assert!(result.is_err(), "expected rejection for {child}");
        }
    }
    assert!(parse_xml("<calcPr/><calcPr/>").is_err());
    assert!(parse_xml(r#"<calcPr calcId="1" calcId="2"/>"#).is_err());
}

#[test]
fn rejects_entity_references_inside_calc_pr() {
    for reference in ["&amp;", "&#x20;"] {
        let child = format!("<calcPr>{reference}</calcPr>");
        assert!(parse_xml(&child).is_err(), "expected rejection for {child}");
    }
}

#[test]
fn rejects_oversized_workbook_xml() {
    assert!(parse(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
}

fn fixture(bytes: &[u8]) -> Properties {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    parse(part.blob()).unwrap().unwrap()
}

#[test]
fn reads_poi_calculation_fixtures() {
    let iterative = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/47889.xlsx"
    )));
    assert_eq!(iterative.calculation_mode(), Mode::Automatic);
    assert!(iterative.iterative_calculation());
    assert_eq!(iterative.iteration_count(), 100);
    assert_eq!(iterative.iteration_delta(), 0.001);

    let no_save = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/58106.xlsx"
    )));
    assert!(!no_save.calculate_on_save());

    let recalculate = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/60289.xlsx"
    )));
    assert!(recalculate.full_calculation_on_load());
}

#[test]
fn reads_libreoffice_calculation_fixtures() {
    let displayed_precision = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/totalsRowShown.xlsx"
    )));
    assert!(!displayed_precision.full_precision());

    let r1c1 = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf134455.xlsx"
    )));
    assert_eq!(r1c1.reference_mode(), ReferenceMode::R1C1);
    assert_eq!(r1c1.iteration_count(), 100);
    assert_eq!(r1c1.iteration_delta(), 0.001);
}

#[test]
fn parses_ordered_strict_features_and_applies_feature_limits() {
    const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const XCALCF: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    const URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
    let xml = format!(
        r#"<s:workbook xmlns:s="{STRICT}" xmlns:f="{XCALCF}"><s:extLst><s:ext uri="{URI}"><f:calcFeatures><f:feature name="A"/><f:feature name="A"/></f:calcFeatures></s:ext></s:extLst></s:workbook>"#,
    );
    let features = parse_features(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(features.len(), 2);
    assert_eq!(features.as_slice()[0].as_str(), "A");
    assert_eq!(features.occurrence_count("A"), 2);

    let limits = Limits::new().with_max_features(1).unwrap();
    assert!(parse_features_with_limits(xml.as_bytes(), &limits).is_err());
}

#[test]
fn rewrite_is_borrowed_on_noop_and_preserves_unowned_bytes() {
    const XCALCF: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    const URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
    let xml = format!(
        r#"<s:workbook xmlns:s="{NS}" xmlns:f="{XCALCF}" xmlns:k="urn:keep"><s:sheets/><s:calcPr calcId="1"/><s:extLst><s:ext uri="urn:keep"><k:opaque exact="yes"/></s:ext><s:ext uri="{URI}"><f:calcFeatures><f:feature name="old"/></f:calcFeatures></s:ext></s:extLst></s:workbook>"#,
    );
    let limits = Limits::default();
    let inspected = inspect(xml.as_bytes(), &limits).unwrap();
    assert!(matches!(
        rewrite(
            &inspected,
            inspected.properties.as_ref(),
            inspected.features.as_ref(),
            &limits,
        )
        .unwrap(),
        Cow::Borrowed(_)
    ));

    let properties = Properties::builder().calculation_id(Some(9)).build();
    let features = Features::new(Feature::new("new&exact").unwrap());
    let changed = rewrite(&inspected, Some(&properties), Some(&features), &limits)
        .unwrap()
        .into_owned();
    let text = String::from_utf8(changed).unwrap();
    assert!(text.contains(r#"<s:calcPr calcId="9"/>"#));
    assert!(text.contains(r#"<k:opaque exact="yes"/>"#));
    assert!(text.contains(r#"<f:feature name="new&amp;exact"/>"#));
}

#[test]
fn rewrite_refuses_mce_projected_calc_pr() {
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let xml = format!(
        r#"<workbook xmlns="{NS}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><calcPr calcId="1"/></mc:Choice><mc:Fallback><calcPr calcId="2"/></mc:Fallback></mc:AlternateContent></workbook>"#,
    );
    let limits = Limits::default();
    let inspected = inspect(xml.as_bytes(), &limits).unwrap();
    let changed = Properties::builder().calculation_id(Some(3)).build();
    assert!(rewrite(&inspected, Some(&changed), None, &limits).is_err());
}

#[test]
fn invalidation_preserves_qualified_future_calc_pr_attributes_exactly() {
    let xml = format!(
        r#"<s:workbook xmlns:s="{NS}"><s:calcPr xmlns:z="urn:future" calcId='9' z:future = 'kept&amp;exact' calcMode="manual"></s:calcPr></s:workbook>"#,
    );
    assert!(parse(xml.as_bytes()).is_err());
    let limits = Limits::default();
    let inspected = inspect(xml.as_bytes(), &limits).unwrap();
    let changed = invalidate_formulas(&inspected, &limits)
        .unwrap()
        .into_owned();
    let text = String::from_utf8(changed).unwrap();
    assert!(text.contains(r#" xmlns:z="urn:future""#));
    assert!(text.contains(r#" z:future = 'kept&amp;exact'"#));
    let reparsed = inspect(text.as_bytes(), &limits).unwrap();
    let properties = reparsed.properties.unwrap();
    assert_eq!(properties.calculation_id(), 0);
    assert!(properties.full_calculation_on_load());
    assert!(properties.force_full_calculation());
    assert!(!properties.calculation_completed());
    assert!(properties.calculate_on_save());
    assert_eq!(properties.calculation_mode(), Mode::Manual);
}

#[test]
fn applies_schema_whitespace_unsigned_and_double_lexicals() {
    let properties = parse_xml(
        r#"<calcPr calcId="  00042  " calcMode="&#x9; autoNoTable &#xA;" fullCalcOnLoad=" true " refMode=" R1C1 " iterateDelta=" -INF "/>"#,
    )
    .unwrap()
    .unwrap();
    assert_eq!(properties.calculation_id(), 42);
    assert_eq!(properties.calculation_mode(), Mode::AutomaticExceptTables);
    assert!(properties.full_calculation_on_load());
    assert_eq!(properties.reference_mode(), ReferenceMode::R1C1);
    assert_eq!(properties.iteration_delta(), f64::NEG_INFINITY);

    for (lexical, expected) in [
        ("INF", f64::INFINITY),
        ("-INF", f64::NEG_INFINITY),
        ("NaN", f64::NAN),
        ("-0", -0.0),
        ("-1.25E+2", -125.0),
    ] {
        let value = parse_xml(&format!(r#"<calcPr iterateDelta=" {lexical} "/>"#))
            .unwrap()
            .unwrap()
            .iteration_delta();
        if expected.is_nan() {
            assert!(value.is_nan());
        } else {
            assert_eq!(value.to_bits(), expected.to_bits());
        }
    }

    let source = format!(r#"<workbook xmlns="{NS}"><calcPr/></workbook>"#);
    let limits = Limits::default();
    for (value, lexical) in [
        (f64::INFINITY, "INF"),
        (f64::NEG_INFINITY, "-INF"),
        (f64::NAN, "NaN"),
        (-0.0, "-0"),
        (-12.5, "-12.5"),
    ] {
        let inspected = inspect(source.as_bytes(), &limits).unwrap();
        let properties = Properties::builder()
            .iteration_delta(Some(value))
            .unwrap()
            .build();
        let output = rewrite(&inspected, Some(&properties), None, &limits)
            .unwrap()
            .into_owned();
        assert!(
            String::from_utf8(output)
                .unwrap()
                .contains(&format!(r#"iterateDelta="{lexical}""#))
        );
    }
}

#[test]
fn unsigned_int_lexicals_are_digits_only_bounded_and_canonically_written() {
    let accepted = parse_xml(
        r#"<calcPr calcId=" 0000000000 " iterateCount="&#x9;4294967295&#xA;" concurrentManualCount="0001"/>"#,
    )
    .unwrap()
    .unwrap();
    assert_eq!(accepted.calculation_id(), 0);
    assert_eq!(accepted.iteration_count(), u32::MAX);
    assert_eq!(accepted.concurrent_manual_count(), Some(1));

    for name in ["calcId", "iterateCount", "concurrentManualCount"] {
        for lexical in ["+0", "+1", "-0", "1 0", "4294967296"] {
            assert!(
                parse_xml(&format!(r#"<calcPr {name}=" {lexical} "/>"#)).is_err(),
                "expected {name}={lexical:?} rejection",
            );
        }
    }

    let source = format!(r#"<workbook xmlns="{NS}"><calcPr calcId="0001"/></workbook>"#);
    let limits = Limits::default();
    let inspected = inspect(source.as_bytes(), &limits).unwrap();
    let properties = Properties::builder()
        .calculation_id(Some(u32::MAX))
        .iteration_count(Some(0))
        .concurrent_manual_count(Some(1))
        .build();
    let output = rewrite(&inspected, Some(&properties), None, &limits)
        .unwrap()
        .into_owned();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"calcId="4294967295""#));
    assert!(text.contains(r#"iterateCount="0""#));
    assert!(text.contains(r#"concurrentManualCount="1""#));
}

#[test]
fn understands_xcalcf_choice_and_collapses_extension_uri_token() {
    const XCALCF: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
    let xml = format!(
        r#"<workbook xmlns="{NS}" xmlns:mc="{MCE}" xmlns:f="{XCALCF}"><mc:AlternateContent><mc:Choice Requires="f"><extLst><ext uri="  {URI} &#xA; "><f:calcFeatures><f:feature name="choice"/></f:calcFeatures></ext></extLst></mc:Choice><mc:Fallback><extLst><ext uri="{URI}"><f:calcFeatures><f:feature name="fallback"/></f:calcFeatures></ext></extLst></mc:Fallback></mc:AlternateContent></workbook>"#,
    );
    let features = parse_features(xml.as_bytes()).unwrap().unwrap();
    assert_eq!(features.as_slice()[0].as_str(), "choice");
}

#[test]
fn rejects_mixed_dialect_owners_and_out_of_order_calc_pr() {
    const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    let mixed =
        format!(r#"<s:workbook xmlns:s="{STRICT}" xmlns:t="{NS}"><t:calcPr/></s:workbook>"#,);
    assert!(parse(mixed.as_bytes()).is_err());
    assert!(parse_xml("<calcPr/><sheets/>").is_err());
}

#[test]
fn serializes_large_feature_collections_with_valid_structure() {
    let mut values = Vec::new();
    values.try_reserve(4096).unwrap();
    for index in 0..4096 {
        values.push(Feature::new(format!("feature-{index}&exact")).unwrap());
    }
    let features = Features::try_from_vec(values).unwrap();
    let xml = format!(r#"<workbook xmlns="{NS}"><sheets/></workbook>"#);
    let limits = Limits::default();
    let inspected = inspect(xml.as_bytes(), &limits).unwrap();
    let output = rewrite(&inspected, None, Some(&features), &limits)
        .unwrap()
        .into_owned();
    let reparsed = parse_features(&output).unwrap().unwrap();
    assert_eq!(reparsed.len(), 4096);
    assert_eq!(reparsed.as_slice()[0].as_str(), "feature-0&exact");
    assert_eq!(reparsed.as_slice()[4095].as_str(), "feature-4095&exact");
}
