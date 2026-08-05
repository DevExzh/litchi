use litchi_odt::list_label_alignment::{Alignment, FollowedBy, Kind, Length, Style, parse};
use litchi_odt::{Builder, Document};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const T: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:t="{T}" xmlns:f="{F}">{x}</o:styles>"#)
}
#[test]
fn parses_aliases_values_and_round_trip() {
    let x = wrap(
        r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="listtab" t:list-tab-stop-position="1.27cm" f:text-indent="-0.635cm" f:margin-left="1.27cm"/></s:list-level-properties></t:list-level-style-number><t:list-level-style-bullet t:level="2"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="space"/></s:list-level-properties></t:list-level-style-bullet><t:list-level-style-image t:level="3"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="nothing"/></s:list-level-properties></t:list-level-style-image></t:list-style>"#,
    );
    let p = parse(&x).unwrap();
    assert_eq!(p.levels.len(), 3);
    let a = &p.get("L", 1).unwrap().alignment;
    assert_eq!(a.label_followed_by, FollowedBy::ListTab);
    assert_eq!(a.text_indent.as_ref().unwrap().as_str(), "-0.635cm");
    let fragment = a.to_xml_fragment().unwrap();
    assert!(fragment.contains(r#"text:label-followed-by="listtab""#));
}
#[test]
fn parses_odfdo_and_libreoffice_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/example.xml");
    let odfdo_parsed = parse(odfdo).unwrap();
    assert!(odfdo_parsed.levels.len() >= 10);
    assert!(
        odfdo_parsed
            .levels
            .iter()
            .any(|x| x.list_style_kind == Kind::Outline)
    );
    let lo = include_bytes!(
        "../../../test-data/libreoffice-core/xmloff/qa/unit/data/differentListStylesInOneList.fodt"
    );
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(lo)).unwrap();
    assert!(!flat.alignments().unwrap().levels.is_empty());
}
#[test]
fn builder_package_and_mutable_round_trip() {
    let mut a = Alignment::new(FollowedBy::Space);
    a.text_indent = Some(Length::new("-1cm").unwrap());
    a.margin_left = Some(Length::new("1cm").unwrap());
    let mut b = Builder::new();
    b.set_numbered_list_level_label_alignment(1, a.clone())
        .unwrap();
    b.add_paragraph("x").unwrap();
    let bytes = b.build().unwrap();
    let p = litchi_odt::generic::Package::from_bytes(bytes.clone()).unwrap();
    assert_eq!(&p.alignments().unwrap().get("L1", 1).unwrap().alignment, &a);
    let mut m =
        litchi_odt::mutable::MutableDocument::from_document(Document::from_bytes(bytes).unwrap())
            .unwrap();
    let replacement = Alignment::new(FollowedBy::Nothing);
    m.set_list_level_label_alignment(&Style::new("L1", 1, replacement.clone()).unwrap())
        .unwrap();
    let p = litchi_odt::generic::Package::from_bytes(m.to_bytes().unwrap()).unwrap();
    assert_eq!(
        &p.alignments().unwrap().get("L1", 1).unwrap().alignment,
        &replacement
    );
}
#[test]
fn rejects_malformed() {
    for x in [
        wrap(
            r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-label-alignment t:label-followed-by="space"/></t:list-level-style-number></t:list-style>"#,
        ),
        wrap(
            r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-width-and-position"><s:list-level-label-alignment t:label-followed-by="space"/></s:list-level-properties></t:list-level-style-number></t:list-style>"#,
        ),
        wrap(
            r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment/></s:list-level-properties></t:list-level-style-number></t:list-style>"#,
        ),
        wrap(
            r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="tab"/></s:list-level-properties></t:list-level-style-number></t:list-style>"#,
        ),
        wrap(
            r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="space" f:margin-left="1em"/></s:list-level-properties></t:list-level-style-number></t:list-style>"#,
        ),
        format!(
            r#"<!DOCTYPE x>{}"#,
            wrap(
                r#"<t:list-style s:name="L"><t:list-level-style-number t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="space"/></s:list-level-properties></t:list-level-style-number></t:list-style>"#
            )
        ),
    ] {
        assert!(parse(&x).is_err(), "accepted {x}");
    }
}
