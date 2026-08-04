//! Regression tests for the drawing-page model and codecs.

use super::{
    Fill, Repeat, Style, StyleProperties, parse_drawing_page_style_properties,
    set_drawing_page_style_properties_xml,
};
const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#;
fn doc(body: &str) -> String {
    format!("{HEAD}{body}</office:automatic-styles></office:document>")
}
#[test]
fn complete_family_round_trips() {
    let xml = doc(
        r##"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="gradient" draw:fill-color="#A0b1C2" draw:secondary-fill-color="#010203" draw:fill-gradient-name="g1" draw:gradient-step-count="00012" draw:fill-hatch-name="h1" draw:fill-hatch-solid="true" draw:fill-image-name="" style:repeat="stretch" draw:fill-image-width="0cm" draw:fill-image-height="-.5%" draw:fill-image-ref-point-x="-12.5%" draw:fill-image-ref-point-y=".5%" draw:fill-image-ref-point="bottom-right" draw:tile-repeat-offset="100.0% vertical" draw:opacity="120%" draw:opacity-name="o1" svg:fill-rule="evenodd" presentation:transition-type="semi-automatic" presentation:transition-style="melt" presentation:transition-speed="fast" smil:type="fade &amp; dissolve" smil:subtype="crossfade" smil:direction="reverse" smil:fadeColor="#aB09fF" presentation:duration="P1Y2M3DT4H5M6.50S" presentation:visibility="hidden" draw:background-size="border" presentation:background-objects-visible="false" presentation:background-visible="true" presentation:display-header="false" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"><presentation:sound xlink:type="simple" xlink:href="https://example.invalid/a&amp;b.ogg" xlink:actuate="onRequest" xlink:show="new" presentation:play-full="true" xml:id="sound_1"/></style:drawing-page-properties></style:style>"##,
    );
    let set = parse_drawing_page_style_properties(&xml).unwrap();
    let p = set.get("dp1").unwrap().properties.as_ref().unwrap();
    assert_eq!(p.gradient_step_count.as_ref().unwrap().as_str(), "00012");
    assert_eq!(
        p.sound.as_ref().unwrap().href,
        "https://example.invalid/a&b.ogg"
    );
    let fragment = p.to_xml_fragment().unwrap();
    assert_eq!(StyleProperties::from_xml_fragment(&fragment).unwrap(), *p)
}
#[test]
fn parses_real_libreoffice_remote_background_without_loading_it() {
    let xml = include_str!(
        "../../../../test-data/libreoffice-core/sd/qa/unit/tiledrendering/data/slide-background-link.fodp"
    );
    let set = parse_drawing_page_style_properties(xml).unwrap();
    let p = set.get("dp1").unwrap().properties.as_ref().unwrap();
    assert_eq!(p.fill, Some(Fill::Bitmap));
    assert_eq!(p.fill_image_name.as_ref().unwrap().as_str(), "remote_bg");
    assert_eq!(p.repeat, Some(Repeat::Stretch))
}
#[test]
fn lossless_replace_insert_and_remove() {
    let original = doc(
        "<!--keep--><style:style style:name=\"a\" style:family=\"drawing-page\"><x:keep xmlns:x=\"urn:keep\"/></style:style><style:style style:name=\"b\" style:family=\"drawing-page\"><style:drawing-page-properties draw:fill=\"none\"/></style:style>",
    );
    let mut a = Style::named(
        "a",
        Some(StyleProperties {
            fill: Some(Fill::Solid),
            ..Default::default()
        }),
    )
    .unwrap();
    let inserted = set_drawing_page_style_properties_xml(&original, &a).unwrap();
    assert!(inserted.contains("<!--keep--><style:style"));
    assert!(inserted.contains("<x:keep xmlns:x=\"urn:keep\"/><style:drawing-page-properties"));
    a.properties = None;
    let removed_a = set_drawing_page_style_properties_xml(&inserted, &a).unwrap();
    assert_eq!(removed_a, original);
    let b = Style::named("b", None).unwrap();
    let removed = set_drawing_page_style_properties_xml(&removed_a, &b).unwrap();
    assert!(!removed.contains("draw:fill=\"none\""));
    assert!(removed.contains("<!--keep-->"))
}
#[test]
fn rejects_malformed_namespaces_lexicals_duplicates_and_children() {
    let cases = [
        r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties presentation:display-header="1"/></style:style>"#,
        r##"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:fill-color="#fff"/></style:style>"##,
        r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:tile-repeat-offset="101% horizontal"/></style:style>"#,
        r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties draw:fill="none" draw:fill="solid"/></style:style>"#,
        r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties><draw:sound/></style:drawing-page-properties></style:style>"#,
        r#"<style:style style:name="a" style:family="drawing-page"><draw:drawing-page-properties/></style:style>"#,
        r#"<style:style style:name="a" style:family="drawing-page"><style:drawing-page-properties><presentation:sound xlink:type="simple" xlink:href="x"><presentation:sound xlink:type="simple" xlink:href="y"/></presentation:sound></style:drawing-page-properties></style:style>"#,
    ];
    for case in cases {
        assert!(
            parse_drawing_page_style_properties(&doc(case)).is_err(),
            "accepted {case}"
        )
    }
}
