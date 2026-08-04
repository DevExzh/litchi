use litchi_odt::{
    FlatOpenDocument, LabelFollowedBy, OdfListLevelPositionMode, OdfOutlineTextAlign,
    parse_outline_styles,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

fn document(outline: &str) -> String {
    format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}" xmlns:f="{FO}" xmlns:x="urn:producer" o:mimetype="application/vnd.oasis.opendocument.text"><o:styles>{outline}</o:styles><o:body><o:text/></o:body></o:document>"#
    )
}

#[test]
fn parses_complete_typed_outline_style_and_extensions() {
    let xml = document(
        r#"<t:outline-style s:name="Outline"><t:outline-level-style t:level="1" t:style-name="Heading_20_1" s:num-format="A" s:num-letter-sync="true" s:num-prefix="(" s:num-suffix=")" t:display-levels="2" t:start-value="3" x:num-list-format="%1%"><s:list-level-properties f:text-align="start" t:space-before="-0.5cm" t:min-label-width="0cm" t:min-label-distance="0.2cm" s:font-name="Liberation Sans" f:width="1cm" f:height="2cm" s:vertical-rel="line" s:vertical-pos="middle" t:list-level-position-and-space-mode="label-alignment" x:producer="kept"><s:list-level-label-alignment t:label-followed-by="listtab" t:list-tab-stop-position="1.2cm" f:text-indent="-0.4cm" f:margin-left="1.2cm"/></s:list-level-properties><s:text-properties f:font-weight="bold" x:feature="kept"/></t:outline-level-style><t:outline-level-style t:level="2" s:num-format=""/></t:outline-style>"#,
    );
    let parsed = parse_outline_styles(&xml).unwrap();
    let style = parsed.get("Outline").unwrap();
    assert_eq!(style.levels.len(), 2);
    let level = style.level(1).unwrap();
    assert_eq!(level.number_format.as_ref().unwrap().as_str(), "A");
    assert_eq!(level.display_levels.as_ref().unwrap().as_str(), "2");
    assert_eq!(level.start_value.as_ref().unwrap().as_str(), "3");
    assert_eq!(level.extensions[0].namespace_uri(), "urn:producer");
    let properties = level.list_level_properties.as_ref().unwrap();
    assert_eq!(properties.text_align, Some(OdfOutlineTextAlign::Start));
    assert_eq!(
        properties.position_mode,
        Some(OdfListLevelPositionMode::LabelAlignment)
    );
    assert_eq!(
        properties
            .label_alignment
            .as_ref()
            .unwrap()
            .label_followed_by,
        LabelFollowedBy::ListTab
    );
    assert_eq!(properties.extensions[0].value(), "kept");
    assert_eq!(level.text_properties.as_ref().unwrap().attributes.len(), 2);
}

#[test]
fn flat_document_exposes_outline_styles_without_interpreting_headings() {
    let xml = document(
        r#"<t:outline-style s:name="Outline"><t:outline-level-style t:level="1" s:num-format="1"/></t:outline-style>"#,
    );
    let document = FlatOpenDocument::from_bytes(xml.into_bytes()).unwrap();
    assert_eq!(document.outline_styles().unwrap().styles.len(), 1);
}

#[test]
fn parses_local_outline_fixture_with_tenth_level_extension() {
    let xml = include_str!("../../../test-data/odf/odt/outline-ten-level-extension.fodt");
    let parsed = parse_outline_styles(xml).unwrap();
    let outline = parsed.get("Outline").unwrap();
    assert_eq!(outline.levels.len(), 10);
    assert_eq!(outline.level(10).unwrap().extensions[0].value(), "%10%");
}

#[test]
fn rejects_invalid_outline_grammar_and_values() {
    let invalid = [
        r#"<t:outline-style><t:outline-level-style t:level="1"/></t:outline-style>"#,
        r#"<t:outline-style s:name="O"/>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="0"/></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1"/><t:outline-level-style t:level="1"/></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1" s:num-format="1" s:num-letter-sync="true"/></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1" t:display-levels="0"/></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1"><s:text-properties/><s:list-level-properties/></t:outline-level-style></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"/></t:outline-level-style></t:outline-style>"#,
        r#"<t:outline-style s:name="O"><t:outline-level-style t:level="1">text</t:outline-level-style></t:outline-style>"#,
    ];
    for outline in invalid {
        assert!(
            parse_outline_styles(&document(outline)).is_err(),
            "{outline}"
        );
    }
    let misplaced = format!(
        r#"<o:document xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:s="{STYLE}"><o:body><o:text><t:outline-style s:name="O"><t:outline-level-style t:level="1"/></t:outline-style></o:text></o:body></o:document>"#
    );
    assert!(parse_outline_styles(&misplaced).is_err());
}
