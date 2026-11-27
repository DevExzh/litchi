//! Regression tests for the chart-style model and codecs.

use super::*;
const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#;
fn doc(body: &str) -> String {
    format!("{HEAD}{body}</office:automatic-styles></office:document>")
}
#[test]
fn complete_family_round_trips() {
    let xml = doc(
        r#"<style:style style:name="ch1" style:family="chart"><style:chart-properties chart:scale-text="true" chart:three-dimensional="false" chart:deep="true" chart:right-angled-axes="false" chart:symbol-type="image" chart:symbol-width=".5cm" chart:symbol-height="0px" chart:sort-by-x-values="true" chart:vertical="false" chart:connect-bars="true" chart:gap-width="-001" chart:overlap="+2" chart:group-bars-per-axis="false" chart:japanese-candle-stick="true" chart:interpolation="b-spline" chart:spline-order="+03" chart:spline-resolution="4" chart:pie-offset="000" chart:angle-offset="344.61" chart:hole-size="-.5%" chart:lines="true" chart:solid-type="pyramid" chart:stacked="false" chart:percentage="true" chart:treat-empty-cells="leave-gap" chart:link-data-style-to-source="false" chart:logarithmic="true" chart:maximum="INF" chart:minimum="-INF" chart:origin="NaN" chart:interval-major="+1.2E-3" chart:interval-minor-divisor="2" chart:tick-marks-major-inner="true" chart:tick-marks-major-outer="false" chart:tick-marks-minor-inner="true" chart:tick-marks-minor-outer="false" chart:reverse-direction="true" chart:display-label="false" chart:text-overlap="true" text:line-break="false" chart:label-arrangement="stagger-even" style:direction="ttb" style:rotation-angle="" chart:data-label-number="value-and-percentage" chart:data-label-text="true" chart:data-label-symbol="false" chart:label-position="near-origin" chart:label-position-negative="avoid-overlap" chart:visible="true" chart:auto-position="false" chart:auto-size="true" chart:mean-value="false" chart:error-category="cell-range" chart:error-percentage="5" chart:error-margin=".1" chart:error-lower-limit="-2" chart:error-upper-limit="2" chart:error-upper-indicator="true" chart:error-lower-indicator="false" chart:series-source="columns" chart:regression-type="power" chart:axis-position="-0.0" chart:axis-label-position="near-axis-other-side" chart:tick-mark-position="at-labels-and-axis" chart:include-hidden-cells="false"><chart:symbol-image xlink:href="Pictures/a&amp;b.svg"/><chart:label-separator><text:p> / <text:span>rich</text:span></text:p></chart:label-separator></style:chart-properties></style:style>"#,
    );
    let set = parse_chart_style_properties(&xml).unwrap();
    let value = set.get("ch1").unwrap().properties.as_ref().unwrap();
    assert_eq!(value.maximum.as_ref().unwrap().as_str(), "INF");
    assert_eq!(
        value.symbol_image.as_ref().unwrap().href,
        "Pictures/a&b.svg"
    );
    let fragment = value.to_xml_fragment().unwrap();
    assert_eq!(
        StyleProperties::from_xml_fragment(&fragment).unwrap(),
        *value
    );
}
#[test]
fn parses_real_libreoffice_standard_style() {
    let fixture = include_str!(
        "../../../../test-data/libreoffice-core/chart2/qa/extras/data/fods/stacked-column-chart.fods"
    );
    let begin = fixture
        .find(r#"<style:style style:name="ch3" style:family="chart">"#)
        .unwrap();
    let end = begin + fixture[begin..].find("</style:style>").unwrap() + "</style:style>".len();
    let set = parse_chart_style_properties(&doc(&fixture[begin..end])).unwrap();
    let value = set.get("ch3").unwrap().properties.as_ref().unwrap();
    assert_eq!(value.stacked, Some(true));
    assert_eq!(value.series_source, Some(SeriesSource::Rows));
    assert_eq!(value.treat_empty_cells, Some(EmptyCellTreatment::LeaveGap));
}
#[test]
fn lossless_replace_insert_remove() {
    let original = doc(
        "<!--keep--><style:style style:name=\"a\" style:family=\"chart\"><x:k xmlns:x=\"urn:k\"/></style:style><style:style style:name=\"b\" style:family=\"chart\"><style:chart-properties chart:lines=\"true\"/></style:style>",
    );
    let mut a = StyleRecord::named(
        "a",
        Some(StyleProperties {
            stacked: Some(true),
            ..Default::default()
        }),
    )
    .unwrap();
    let inserted = set_chart_style_properties_xml(&original, &a).unwrap();
    assert!(inserted.contains("<x:k xmlns:x=\"urn:k\"/><style:chart-properties"));
    a.properties = None;
    let restored = set_chart_style_properties_xml(&inserted, &a).unwrap();
    assert_eq!(restored, original);
    let b = StyleRecord::named("b", None).unwrap();
    let removed = set_chart_style_properties_xml(&restored, &b).unwrap();
    assert!(!removed.contains("chart:lines=\"true\""));
    assert!(removed.contains("<!--keep-->"));
}
#[test]
fn rejects_malformed_lexicals_namespaces_placement_and_children() {
    let cases = [
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:stacked="1"/></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:spline-order="0"/></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:symbol-type="named-symbol"/></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:symbol-type="image"><chart:symbol-image/></style:chart-properties></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:lines="true" chart:lines="false"/></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><chart:chart-properties/></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties><chart:label-separator/></style:chart-properties></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties><chart:label-separator><chart:p/></chart:label-separator></style:chart-properties></style:style>"#,
        r#"<style:style style:name="a" style:family="chart"><style:chart-properties lo:extension="1" xmlns:lo="urn:extension"/></style:style>"#,
    ];
    for case in cases {
        assert!(
            parse_chart_style_properties(&doc(case)).is_err(),
            "accepted {case}"
        );
    }
}
