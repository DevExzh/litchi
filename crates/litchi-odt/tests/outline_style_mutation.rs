use litchi_odt::outline_style::{
    parse_outline_styles, remove_outline_style_xml, set_outline_style_xml,
};

const XML: &str = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:x="urn:producer"><o:styles><!--keep--><t:outline-style s:name="Outline"><t:outline-level-style t:level="1" s:num-format="A" s:num-letter-sync="true" x:num-list-format="%1%"><s:text-properties f:font-weight="bold"/></t:outline-level-style></t:outline-style><s:style s:name="keep" s:family="paragraph"/></o:styles></o:document-styles>"#;

#[test]
fn canonical_serialization_retains_typed_and_extension_metadata() {
    let style = parse_outline_styles(XML).unwrap().styles.remove(0);
    let serialized = style.to_xml().unwrap();
    assert!(serialized.contains("xmlns:ext0=\"urn:producer\""));
    assert!(serialized.contains("ext0:num-list-format=\"%1%\""));
    assert!(serialized.contains("fo:font-weight=\"bold\""));
    let wrapped = format!(
        r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:styles>{serialized}</o:styles></o:document-styles>"#
    );
    assert_eq!(parse_outline_styles(&wrapped).unwrap().styles, vec![style]);
}

#[test]
fn replace_insert_and_remove_preserve_unrelated_xml() {
    let mut replacement = parse_outline_styles(XML).unwrap().styles.remove(0);
    replacement.name = "Replacement".to_string();
    replacement.levels[0].number_prefix = Some("(".to_string());

    let empty = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:styles/></o:document-styles>"#;
    let (inserted, old) = set_outline_style_xml(empty, &replacement).unwrap();
    assert!(old.is_none());
    assert_eq!(
        parse_outline_styles(&inserted).unwrap().get("Replacement"),
        Some(&replacement)
    );

    let original = parse_outline_styles(XML).unwrap().styles.remove(0);
    let mut changed = original.clone();
    changed.levels[0].number_suffix = Some(")".to_string());
    let (updated, old) = set_outline_style_xml(XML, &changed).unwrap();
    assert_eq!(old, Some(original.clone()));
    assert!(updated.contains("<!--keep-->"));
    assert!(updated.contains("<s:style s:name=\"keep\" s:family=\"paragraph\"/>"));

    let (removed, old) = remove_outline_style_xml(&updated, "Outline").unwrap();
    assert_eq!(old, Some(changed));
    assert!(parse_outline_styles(&removed).unwrap().styles.is_empty());
    assert!(removed.contains("<!--keep-->"));
    assert!(removed.contains("<s:style s:name=\"keep\" s:family=\"paragraph\"/>"));

    let (unchanged, old) = remove_outline_style_xml(&removed, "missing").unwrap();
    assert_eq!(old, None);
    assert_eq!(unchanged, removed);
}
