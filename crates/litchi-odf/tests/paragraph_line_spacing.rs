use litchi_odf::{
    FlatOpenDocument, LineHeight, LineHeightPercent, LineSpacingLength, OdfNonNegativeLength,
    ParagraphLineSpacing, TextAlignLast, TextAutospace, parse_paragraph_style_line_spacings,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}">{x}</o:styles>"#)
}

fn full_properties() -> &'static str {
    r#"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties f:line-height="115%" s:line-spacing="0.2cm" s:line-height-at-least="0.5cm" s:font-independent-line-spacing="true" s:text-autospace="ideograph-alpha" s:justify-single-word="false" s:auto-text-indent="true" s:snap-to-layout-grid="false" s:tab-stop-distance="1.251cm" f:text-align-last="center" f:margin-left="1cm"/></s:style>"#
}

#[test]
fn parses_all_line_spacing_attributes() {
    let set = parse_paragraph_style_line_spacings(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(
        p.line_height,
        Some(LineHeight::Percent(LineHeightPercent::new("115%").unwrap()))
    );
    assert_eq!(
        p.line_spacing,
        Some(LineSpacingLength::new("0.2cm").unwrap())
    );
    assert_eq!(
        p.line_height_at_least,
        Some(OdfNonNegativeLength::new("0.5cm").unwrap())
    );
    assert_eq!(p.font_independent_line_spacing, Some(true));
    assert_eq!(p.text_autospace, Some(TextAutospace::IdeographAlpha));
    assert_eq!(p.justify_single_word, Some(false));
    assert_eq!(p.auto_text_indent, Some(true));
    assert_eq!(p.snap_to_layout_grid, Some(false));
    assert_eq!(
        p.tab_stop_distance,
        Some(OdfNonNegativeLength::new("1.251cm").unwrap())
    );
    assert_eq!(p.text_align_last, Some(TextAlignLast::Center));
}

#[test]
fn parses_normal_and_length_line_heights() {
    let x = wrap(
        r#"<s:style s:name="N" s:family="paragraph"><s:paragraph-properties f:line-height="normal"/></s:style><s:style s:name="L" s:family="paragraph"><s:paragraph-properties f:line-height="0.6cm"/></s:style>"#,
    );
    let set = parse_paragraph_style_line_spacings(&x).unwrap();
    assert_eq!(
        set.get("N").unwrap().properties.as_ref().unwrap().line_height,
        Some(LineHeight::Normal)
    );
    assert_eq!(
        set.get("L").unwrap().properties.as_ref().unwrap().line_height,
        Some(LineHeight::Length(
            OdfNonNegativeLength::new("0.6cm").unwrap()
        ))
    );
}

#[test]
fn resolves_through_parent_and_default_style() {
    let x = wrap(
        r#"<s:default-style s:family="paragraph"><s:paragraph-properties s:tab-stop-distance="1cm"/></s:default-style><s:style s:name="Base" s:family="paragraph"><s:paragraph-properties f:text-align-last="justify"/></s:style><s:style s:name="Child" s:family="paragraph" s:parent-style-name="Base"/>"#,
    );
    let set = parse_paragraph_style_line_spacings(&x).unwrap();
    let resolved = set.resolved("Child").unwrap().unwrap();
    assert_eq!(resolved.text_align_last, Some(TextAlignLast::Justify));
    let default = set.resolved("Missing").unwrap().unwrap();
    assert_eq!(
        default.tab_stop_distance,
        Some(OdfNonNegativeLength::new("1cm").unwrap())
    );
}

#[test]
fn fragment_round_trip() {
    let mut p = ParagraphLineSpacing::new();
    p.line_height = Some(LineHeight::Normal);
    p.text_autospace = Some(TextAutospace::None);
    p.justify_single_word = Some(true);
    p.text_align_last = Some(TextAlignLast::Start);
    p.tab_stop_distance = Some(OdfNonNegativeLength::new("2cm").unwrap());
    let fragment = p.to_xml_fragment().unwrap();
    assert!(fragment.contains(r#"fo:line-height="normal""#));
    assert!(fragment.contains(r#"style:text-autospace="none""#));
    assert!(fragment.contains(r#"fo:text-align-last="start""#));

    let style = format!(r#"<s:style s:name="RT" s:family="paragraph">{fragment}</s:style>"#);
    let set = parse_paragraph_style_line_spacings(&wrap(&style)).unwrap();
    assert_eq!(set.get("RT").unwrap().properties.as_ref(), Some(&p));
}

#[test]
fn styles_without_line_spacing_yield_no_properties() {
    let x = wrap(
        r#"<s:style s:name="Plain" s:family="paragraph"><s:paragraph-properties f:margin-left="1cm"/></s:style><s:style s:name="T" s:family="text"><s:paragraph-properties f:line-height="normal"/></s:style>"#,
    );
    let set = parse_paragraph_style_line_spacings(&x).unwrap();
    assert!(set.get("Plain").unwrap().properties.is_none());
    assert!(set.get("T").is_none());
}

#[test]
fn parses_flat_fixtures() {
    let odfdo = include_str!("../../../test-data/odfdo/tests/samples/example.xml");
    let set = parse_paragraph_style_line_spacings(odfdo).unwrap();
    assert!(
        set.styles
            .iter()
            .filter_map(|style| style.properties.as_ref())
            .any(|p| p.text_autospace == Some(TextAutospace::IdeographAlpha))
    );
    assert!(
        set.styles
            .iter()
            .filter_map(|style| style.properties.as_ref())
            .any(|p| p.tab_stop_distance.is_some())
    );

    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/sw/qa/uibase/shells/data/protectedLinkCopy.fodt");
    let flat = FlatOpenDocument::from_reader(Cursor::new(bytes)).unwrap();
    let set = flat.paragraph_style_line_spacings().unwrap();
    assert!(
        set.styles
            .iter()
            .filter_map(|style| style.properties.as_ref())
            .any(|p| matches!(p.line_height, Some(LineHeight::Percent(_))))
    );
    assert!(
        set.styles
            .iter()
            .filter_map(|style| style.properties.as_ref())
            .any(|p| p.auto_text_indent == Some(false))
    );
}

#[test]
fn rejects_invalid_values() {
    // Bad enum value.
    let x = wrap(
        r#"<s:style s:name="E1" s:family="paragraph"><s:paragraph-properties s:text-autospace="wide"/></s:style>"#,
    );
    assert!(parse_paragraph_style_line_spacings(&x).is_err());
    // Bad boolean.
    let x = wrap(
        r#"<s:style s:name="E2" s:family="paragraph"><s:paragraph-properties s:auto-text-indent="yes"/></s:style>"#,
    );
    assert!(parse_paragraph_style_line_spacings(&x).is_err());
    // Bad length.
    let x = wrap(
        r#"<s:style s:name="E3" s:family="paragraph"><s:paragraph-properties s:line-spacing="tall"/></s:style>"#,
    );
    assert!(parse_paragraph_style_line_spacings(&x).is_err());
    // Negative non-negative length.
    let x = wrap(
        r#"<s:style s:name="E4" s:family="paragraph"><s:paragraph-properties s:tab-stop-distance="-1cm"/></s:style>"#,
    );
    assert!(parse_paragraph_style_line_spacings(&x).is_err());
    // Duplicate paragraph-properties.
    let x = wrap(
        r#"<s:style s:name="E5" s:family="paragraph"><s:paragraph-properties f:line-height="normal"/><s:paragraph-properties f:line-height="normal"/></s:style>"#,
    );
    assert!(parse_paragraph_style_line_spacings(&x).is_err());
}

#[test]
fn rejects_invalid_newtypes() {
    assert!(LineSpacingLength::new("1").is_err());
    assert!(LineSpacingLength::new("cm").is_err());
    assert!(LineSpacingLength::new("1cmx").is_err());
    assert!(LineHeightPercent::new("115").is_err());
    assert!(LineHeightPercent::new("%").is_err());
}
