use litchi_odt::{
    DocumentBuilder, FlatOpenDocument, HyphenationKeep, HyphenationLadder, Keep, LineBreak,
    OpenDocumentPackage, ParagraphFlowProperties, ParagraphStyleFlow, PunctuationWrap,
    parse_paragraph_style_flows,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}">{x}</o:styles>"#)
}
#[test]
fn parses_aliases_values_and_round_trip() {
    let x = wrap(
        r#"<s:default-style s:family="paragraph"><s:paragraph-properties f:keep-together="auto" f:widows="2"/></s:default-style><s:style s:name="Flow" s:family="paragraph"><s:paragraph-properties f:keep-together="always" f:keep-with-next="always" f:widows="3" f:orphans="4" f:hyphenation-keep="page" f:hyphenation-ladder-count="1" s:line-break="strict" s:punctuation-wrap="hanging"/></s:style>"#,
    );
    let p = parse_paragraph_style_flows(&x).unwrap();
    let f = p.get("Flow").unwrap().properties.as_ref().unwrap();
    assert_eq!(f.keep_together, Some(Keep::Always));
    assert_eq!(
        f.hyphenation_ladder_count,
        Some(HyphenationLadder::Lines(1))
    );
    let fragment = p.get("Flow").unwrap().to_xml_fragment().unwrap();
    assert_eq!(
        parse_paragraph_style_flows(&wrap(&fragment))
            .unwrap()
            .get("Flow"),
        p.get("Flow")
    );
}
#[test]
fn parses_real_odfdo_and_libreoffice() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/example.xml");
    assert!(
        !parse_paragraph_style_flows(odfdo)
            .unwrap()
            .styles
            .is_empty()
    );
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/xmloff/qa/unit/data/scale-width-redline.fodt"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(lo)).unwrap();
    assert!(!flat.paragraph_style_flows().unwrap().styles.is_empty());
}
#[test]
fn builder_package_round_trip() {
    let p = ParagraphFlowProperties {
        keep_together: Some(Keep::Always),
        keep_with_next: Some(Keep::Auto),
        widows: Some(2),
        orphans: Some(3),
        hyphenation_keep: Some(HyphenationKeep::Auto),
        hyphenation_ladder_count: Some(HyphenationLadder::NoLimit),
        line_break: Some(LineBreak::Normal),
        punctuation_wrap: Some(PunctuationWrap::Simple),
    };
    let style = ParagraphStyleFlow::named("Flow", Some(p)).unwrap();
    let mut b = DocumentBuilder::new();
    b.add_paragraph_flow_style(style.clone()).unwrap();
    b.add_paragraph("x").unwrap();
    let package = OpenDocumentPackage::from_bytes(b.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_flows().unwrap().get("Flow"),
        Some(&style)
    );
}
#[test]
fn rejects_malformed() {
    for x in [
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties f:keep-together="true"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties f:widows="-1"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties f:orphans="1000001"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties f:hyphenation-keep="column"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties s:line-break="loose"/></s:style>"#,
        ),
        wrap(
            r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties/><s:paragraph-properties f:orphans="2"/></s:style>"#,
        ),
        format!(
            "<!DOCTYPE x>{}",
            wrap(
                r#"<s:style s:name="X" s:family="paragraph"><s:paragraph-properties f:orphans="2"/></s:style>"#
            )
        ),
    ] {
        assert!(parse_paragraph_style_flows(&x).is_err(), "accepted {x}");
    }
}
