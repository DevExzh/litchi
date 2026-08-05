//! Focused regression tests for the ODF ruby owner.

use super::*;

const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:styles>"#;
fn styles(body: &str) -> String {
    format!(
        "{HEAD}{body}</office:styles><office:body><office:text><text:p/></office:text></office:body></office:document>"
    )
}
#[test]
fn exhaustive_properties_round_trip() {
    for position in Position::ALL {
        for alignment in Alignment::ALL {
            let style = Style::new(
                format!("s{}_{}", position.as_str(), alignment.as_str()),
                Some(Properties {
                    position: Some(position),
                    alignment: Some(alignment),
                }),
            )
            .unwrap();
            let xml = styles(&style.to_xml_fragment().unwrap());
            assert_eq!(parse_ruby_styles(&xml).unwrap().styles[0], style);
        }
    }
}
#[test]
fn mixed_inline_round_trip() {
    let base = Base::from_xml_fragment(r#" A <text:span text:style-name="Em"> B </text:span><text:a xlink:type="simple" xlink:href="https://example.invalid/">一日</text:a> "#).unwrap();
    let ruby = Annotation::new(
        Some("Ru1".into()),
        base,
        "ついたち",
        Some("RubyText".into()),
    )
    .unwrap();
    let parsed = Annotation::from_xml_fragment(&ruby.to_xml_fragment().unwrap()).unwrap();
    assert_eq!(parsed, ruby);
}
#[test]
fn parses_real_libreoffice_fixture() {
    let xml = include_str!("../../../../test-data/odf/odt/ruby-hyperlink.fodt");
    let styles = parse_ruby_styles(xml).unwrap();
    assert_eq!(
        styles.styles[0].properties.as_ref().unwrap().alignment,
        Some(Alignment::Left)
    );
    let annotations = parse_ruby_annotations(xml).unwrap();
    assert_eq!(annotations.annotations[0].text, "ついたち");
    assert!(annotations.annotations[0].base.xml().contains("xlink:href"));
}
#[test]
fn rejects_malformed_order_namespace_and_lexicals() {
    for body in [
        r#"<text:ruby><text:ruby-text>x</text:ruby-text><text:ruby-base>X</text:ruby-base></text:ruby>"#,
        r#"<text:ruby><text:ruby-base>X</text:ruby-base><text:ruby-text><text:span>x</text:span></text:ruby-text></text:ruby>"#,
        r#"<text:ruby text:style-name="1bad"><text:ruby-base>X</text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby>"#,
        r#"<text:ruby><text:ruby-base><text:unknown/></text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby>"#,
    ] {
        let xml = format!(
            r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</text:p>"#
        );
        assert!(parse_ruby_annotations(&xml).is_err(), "accepted {body}");
    }
    let wrong = styles(
        r#"<style:style style:name="r" style:family="ruby"><style:ruby-properties style:ruby-align="justify"/></style:style>"#,
    );
    assert!(parse_ruby_styles(&wrong).is_err());
    let duplicate = styles(
        r#"<style:style style:name="r" style:family="ruby"><style:ruby-properties style:ruby-align="left"/><style:ruby-properties/></style:style>"#,
    );
    assert!(parse_ruby_styles(&duplicate).is_err());
}
#[test]
fn enforces_caps() {
    assert!(Base::from_text(&"x".repeat(MAX_BASE + 1)).is_err());
    let mut xml =
        String::from(r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#);
    for _ in 0..=MAX_DEPTH {
        xml.push_str("<text:span>");
    }
    for _ in 0..=MAX_DEPTH {
        xml.push_str("</text:span>");
    }
    xml.push_str("</text:p>");
    assert!(parse_ruby_annotations(&xml).is_err());
    let unit = "<text:ruby><text:ruby-base/><text:ruby-text/></text:ruby>";
    let sequential = format!(
        r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{}</text:p>"#,
        unit.repeat(MAX_RUBIES + 1)
    );
    assert!(parse_ruby_annotations(&sequential).is_err());
}
#[test]
fn ruby_style_namespace_rejection_is_attribute_order_independent() {
    let xml = styles(
        r#"<style:style xmlns:x="urn:wrong" x:foreign="first" style:name="Ru" style:family="ruby"/>"#,
    );
    assert!(parse_ruby_styles(&xml).is_err());
}
#[test]
fn lossless_style_and_inline_mutation() {
    let original = styles("<!--keep--><style:style style:name=\"other\" style:family=\"text\"/>");
    let style = Style::new(
        "Ru1",
        Some(Properties {
            position: Some(Position::Above),
            alignment: Some(Alignment::Center),
        }),
    )
    .unwrap();
    let inserted = set_ruby_style_xml(&original, &style).unwrap();
    assert!(inserted.contains("<!--keep--><style:style style:name=\"other\""));
    assert_eq!(remove_ruby_style_xml(&inserted, "Ru1").unwrap(), original);
    let ruby = Annotation::new(None, Base::from_text("base").unwrap(), "ruby", None).unwrap();
    let paragraph =
        r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">before</text:p>"#;
    let with = insert_ruby_annotation_xml(paragraph, 0, &ruby).unwrap();
    let replacement = Annotation::new(None, Base::from_text("B").unwrap(), "R", None).unwrap();
    let replaced = replace_ruby_annotation_xml(&with, 0, &replacement).unwrap();
    assert!(replaced.contains(">B</text:ruby-base>"));
    assert_eq!(remove_ruby_annotation_xml(&replaced, 0).unwrap(), paragraph);
}
#[test]
fn wraps_one_paragraph_text_node_range_without_crossing_markup() {
    let paragraph = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">前 <text:span text:style-name="Em">漢字</text:span> 後</text:p>"#;
    let ruby = Annotation::new(None, Base::from_text("字").unwrap(), "じ", None).unwrap();
    let start = "前 ".len() + "漢".len();
    let wrapped = wrap_ruby_annotation_xml(paragraph, 0, start..start + "字".len(), &ruby).unwrap();
    assert!(wrapped.contains("<text:span text:style-name=\"Em\">漢<text:ruby"));
    assert_eq!(
        parse_ruby_annotations(&wrapped).unwrap().annotations,
        vec![ruby.clone()]
    );

    let entity_paragraph =
        r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A &amp; B</text:p>"#;
    let entity = Annotation::new(None, Base::from_text("&").unwrap(), "and", None).unwrap();
    let entity_wrapped = wrap_ruby_annotation_xml(entity_paragraph, 0, 2..3, &entity).unwrap();
    assert_eq!(
        parse_ruby_annotations(&entity_wrapped).unwrap().annotations,
        vec![entity]
    );

    assert!(wrap_ruby_annotation_xml(paragraph, 0, 0..start + 1, &ruby).is_err());
    assert!(wrap_ruby_annotation_xml(paragraph, 0, start + 1..start + "字".len(), &ruby).is_err());
    let wrong_base = Annotation::new(None, Base::from_text("語").unwrap(), "ご", None).unwrap();
    assert!(
        wrap_ruby_annotation_xml(paragraph, 0, start..start + "字".len(), &wrong_base).is_err()
    );
}
#[test]
fn wraps_adjacent_text_cdata_and_entity_nodes() {
    let paragraph = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A<![CDATA[漢]]>&amp;字Z</text:p>"#;
    let base = String::from("漢&字");
    let ruby = Annotation::new(None, Base::from_text(&base).unwrap(), "かんじ", None).unwrap();
    let wrapped = wrap_ruby_annotation_xml(paragraph, 0, 1..1 + base.len(), &ruby).unwrap();

    assert!(wrapped.starts_with(
        r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A<text:ruby"#
    ));
    assert!(wrapped.ends_with("</text:ruby>Z</text:p>"));
    assert_eq!(
        parse_ruby_annotations(&wrapped).unwrap().annotations,
        vec![ruby]
    );
}

#[test]
fn plain_base_range_still_refuses_inline_markup_boundaries() {
    let paragraph = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A<![CDATA[B]]><text:span>C</text:span>D</text:p>"#;
    let ruby = Annotation::new(None, Base::from_text("BCD").unwrap(), "base", None).unwrap();
    assert!(wrap_ruby_annotation_xml(paragraph, 0, 1..4, &ruby).is_err());
}

#[test]
fn wraps_balanced_inline_markup_as_a_structured_base() {
    let paragraph = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A<text:span text:style-name="Em">漢</text:span><text:span text:style-name="Strong">字</text:span>Z</text:p>"#;
    let base = Base::from_xml_fragment(
            r#"<text:span text:style-name="Em">漢</text:span><text:span text:style-name="Strong">字</text:span>"#,
        )
        .unwrap();
    let ruby = Annotation::new(None, base, "かんじ", None).unwrap();
    let wrapped = wrap_ruby_annotation_xml(paragraph, 0, 1..1 + "漢字".len(), &ruby).unwrap();

    assert!(
        wrapped.contains(
            r#">A<text:ruby xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#
        )
    );
    assert!(wrapped.ends_with("</text:ruby>Z</text:p>"));
    assert_eq!(
        parse_ruby_annotations(&wrapped).unwrap().annotations,
        vec![ruby]
    );
}

#[test]
fn structural_range_refuses_partial_ancestors_and_existing_ruby() {
    let partial = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><text:span text:style-name="Em">AB</text:span>C</text:p>"#;
    let partial_base =
        Base::from_xml_fragment(r#"<text:span text:style-name="Em">B</text:span>C"#).unwrap();
    let partial_ruby = Annotation::new(None, partial_base, "partial", None).unwrap();
    assert!(wrap_ruby_annotation_xml(partial, 0, 1..3, &partial_ruby).is_err());

    let existing = Annotation::new(None, Base::from_text("X").unwrap(), "existing", None)
        .unwrap()
        .to_xml_fragment()
        .unwrap();
    let paragraph = format!(
        r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">A{existing}B</text:p>"#
    );
    let crossing_base = Base::from_xml_fragment(&format!("A{existing}B")).unwrap();
    let crossing = Annotation::new(None, crossing_base, "outer", None).unwrap();
    assert!(wrap_ruby_annotation_xml(&paragraph, 0, 0..2, &crossing).is_err());
}
