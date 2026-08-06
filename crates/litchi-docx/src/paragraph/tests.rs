//! Behavioral tests for the paragraph and run facade.

use crate::UnderlineStyle;
use crate::color::Theme;
use litchi_core::VerticalPosition;

use super::*;

#[test]
fn test_run_text_extraction() {
    let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Hello, World!</w:t>
        </w:r>"#;

    let run = Run::new(xml.to_vec());
    let text = run.text().unwrap();
    assert_eq!(text, "Hello, World!");
}

#[test]
fn extracts_decoded_word_text_and_special_characters() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:t xml:space="preserve">  A &amp; B &lt; C &#x1F600;  </w:t></w:r>
            <w:r><w:tab/><w:t/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:t>tail</w:t></w:r>
        </w:p>"#;
    let paragraph = Paragraph::new(xml.to_vec());
    assert_eq!(
        paragraph.text().unwrap(),
        "  A & B < C 😀  \t\n\n‑\u{00ad}tail"
    );
    let runs = paragraph.runs().unwrap();
    assert_eq!(runs[0].text().unwrap(), "  A & B < C 😀  ");
    assert_eq!(runs[1].text().unwrap(), "\t\n\n‑\u{00ad}tail");
}

#[test]
fn runs_resolve_namespace_aliases_and_ignore_lookalikes() {
    let xml = br#"<wp:p xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml">
            <false:r><false:t>ignored outer</false:t></false:r>
            <wp:r><false:t>ignored inner</false:t><false:tab/><false:br/><wp:t>kept</wp:t><wp:tab/></wp:r>
            <wp:r/>
        </wp:p>"#;

    let paragraph = Paragraph::new(xml.to_vec());
    assert_eq!(paragraph.text().unwrap(), "kept\t");
    let runs = paragraph.runs().unwrap();
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[0].text().unwrap(), "kept\t");
    assert_eq!(runs[1].text().unwrap(), "");
}

#[test]
fn runs_accept_the_strict_wordprocessingml_namespace() {
    let xml = br#"<s:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main">
            <s:r><s:t>strict</s:t></s:r>
        </s:p>"#;

    let paragraph = Paragraph::new(xml.to_vec());
    assert_eq!(paragraph.text().unwrap(), "strict");
    let runs = paragraph.runs().unwrap();
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text().unwrap(), "strict");
}

#[test]
fn runs_accept_fragments_with_an_inherited_namespace_binding() {
    let xml = br#"<wp:p><wp:r><wp:t>inherited</wp:t></wp:r></wp:p>"#;

    let paragraph = Paragraph::new(xml.to_vec());
    assert_eq!(paragraph.text().unwrap(), "inherited");
    let runs = paragraph.runs().unwrap();
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text().unwrap(), "inherited");
}

#[test]
fn reads_nested_smart_tags_and_their_typed_metadata() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:false="urn:not-wordprocessingml">
            <w:smartTag w:uri="urn:contacts" w:element="person">
                <w:smartTagPr>
                    <w:attr w:uri="urn:meta" w:name="kind" w:val="friend &amp; peer"/>
                </w:smartTagPr>
                <w:r><w:t>A &amp; </w:t></w:r>
                <w:smartTag w:element="givenName">
                    <w:smartTagPr><w:attr w:name="language" w:val="en"/></w:smartTagPr>
                    <w:r><w:t>Bob</w:t></w:r>
                </w:smartTag>
            </w:smartTag>
            <false:smartTag false:element="ignored"><w:r><w:t>not a tag</w:t></w:r></false:smartTag>
            <w:smartTag w:element="empty"/>
        </w:p>"#;

    let paragraph = Paragraph::new(xml.to_vec());
    let tags = paragraph.smart_tags().unwrap();
    assert_eq!(tags.len(), 3);

    assert_eq!(tags[0].uri.as_deref(), Some("urn:contacts"));
    assert_eq!(tags[0].element, "person");
    assert_eq!(tags[0].attributes.len(), 1);
    assert_eq!(tags[0].attributes[0].uri.as_deref(), Some("urn:meta"));
    assert_eq!(tags[0].attributes[0].name, "kind");
    assert_eq!(tags[0].attributes[0].value, "friend & peer");
    assert_eq!(tags[0].text().unwrap(), "A & Bob");

    assert_eq!(tags[1].element, "givenName");
    assert_eq!(tags[1].attributes[0].name, "language");
    assert_eq!(tags[1].text().unwrap(), "Bob");
    assert_eq!(tags[2].element, "empty");
    assert_eq!(tags[2].text().unwrap(), "");

    assert_eq!(paragraph.runs().unwrap().len(), 3);
}

#[test]
fn smart_tags_require_schema_mandated_attributes() {
    let missing_element = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTag><w:r><w:t>invalid</w:t></w:r></w:smartTag></w:p>"#
            .to_vec(),
    );
    assert!(missing_element.smart_tags().is_err());

    let missing_property_value = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:smartTag w:element="person"><w:smartTagPr><w:attr w:name="kind"/></w:smartTagPr></w:smartTag></w:p>"#
            .to_vec(),
    );
    assert!(missing_property_value.smart_tags().is_err());
}

#[test]
fn runs_reject_unterminated_run_xml() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>truncated</w:t>"#;
    let paragraph = Paragraph::new(xml.to_vec());
    assert!(paragraph.text().is_err());
    assert!(paragraph.runs().is_err());
}

#[test]
fn optimized_run_extraction_matches_text_and_reads_qualified_properties() {
    let run = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:rPr><w:b w:val="0"/><w:i w:val="on"/><w:strike/><w:u w:val="dashLongHeavy"/><w:vertAlign w:val="superscript"/></w:rPr>
                <w:t xml:space="preserve"> A &amp; <![CDATA[B < C]]> &#x1F600; </w:t><w:t/>
                <w:tab/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:t>tail</w:t>
            </w:r>"#
            .to_vec(),
    );

    let expected_text = " A & B < C 😀 \t\n\n‑\u{00ad}tail";
    let (text, properties) = run.get_text_and_properties().unwrap();
    assert_eq!(text, expected_text);
    assert_eq!(text, run.text().unwrap());
    assert_eq!(properties.bold, Some(false));
    assert_eq!(properties.italic, Some(true));
    assert_eq!(properties.strikethrough, Some(true));
    assert_eq!(properties.underline, Some(UnderlineStyle::DashLongHeavy));
    assert_eq!(
        properties.vertical_position,
        Some(VerticalPosition::Superscript)
    );

    let properties_only = run.get_properties().unwrap();
    assert_eq!(properties_only.bold, properties.bold);
    assert_eq!(properties_only.italic, properties.italic);
    assert_eq!(properties_only.strikethrough, properties.strikethrough);
    assert_eq!(properties_only.underline, properties.underline);
    assert_eq!(
        properties_only.vertical_position,
        properties.vertical_position
    );
    assert_eq!(run.bold().unwrap(), Some(false));
    assert_eq!(run.italic().unwrap(), Some(true));
    assert_eq!(
        run.underline_style().unwrap(),
        Some(UnderlineStyle::DashLongHeavy)
    );
    assert_eq!(
        run.vertical_position().unwrap(),
        properties.vertical_position
    );
}

#[test]
fn reads_every_wordprocessingml_underline_pattern() {
    let patterns = [
        ("none", UnderlineStyle::None),
        ("single", UnderlineStyle::Single),
        ("words", UnderlineStyle::Words),
        ("double", UnderlineStyle::Double),
        ("thick", UnderlineStyle::Thick),
        ("dotted", UnderlineStyle::Dotted),
        ("dottedHeavy", UnderlineStyle::DottedHeavy),
        ("dash", UnderlineStyle::Dashed),
        ("dashedHeavy", UnderlineStyle::DashedHeavy),
        ("dashLong", UnderlineStyle::DashLong),
        ("dashLongHeavy", UnderlineStyle::DashLongHeavy),
        ("dotDash", UnderlineStyle::DotDash),
        ("dashDotHeavy", UnderlineStyle::DashDotHeavy),
        ("dotDotDash", UnderlineStyle::DotDotDash),
        ("dashDotDotHeavy", UnderlineStyle::DashDotDotHeavy),
        ("wave", UnderlineStyle::Wave),
        ("wavyHeavy", UnderlineStyle::WavyHeavy),
        ("wavyDouble", UnderlineStyle::WavyDouble),
    ];

    for (value, expected) in patterns {
        let run = Run::new(
            format!(
                r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:val="{value}"/></w:rPr></w:r>"#
            )
            .into_bytes(),
        );
        assert_eq!(run.underline_style().unwrap(), Some(expected));
        assert_eq!(run.get_properties().unwrap().underline, Some(expected));
        assert_eq!(
            run.underline().unwrap(),
            Some(expected != UnderlineStyle::None)
        );
    }

    let implicit_single = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u/></w:rPr></w:r>"#
            .to_vec(),
    );
    assert_eq!(
        implicit_single.underline_style().unwrap(),
        Some(UnderlineStyle::Single)
    );
    let inherited = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr/></w:r>"#
            .to_vec(),
    );
    assert_eq!(inherited.underline().unwrap(), None);
}

#[test]
fn reads_complete_underline_metadata_namespace_aware() {
    let strict = Run::new(
        br#"<s:r xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml"><s:rPr><false:u false:val="double"/><s:u s:val="wavyDouble" s:color="A0b1C2" s:themeColor="accent4" s:themeTint="0a" s:themeShade="FF"/></s:rPr></s:r>"#
            .to_vec(),
    );
    assert_eq!(
        strict.underline_formatting().unwrap(),
        Some(RunUnderline {
            style: UnderlineStyle::WavyDouble,
            color: Some(RunUnderlineColor::Rgb([0xA0, 0xB1, 0xC2])),
            theme_color: Some(Theme::Accent4),
            theme_tint: Some(0x0A),
            theme_shade: Some(0xFF),
        })
    );

    let inherited =
        Run::new(br#"<q:r><q:rPr><q:u q:val="words" q:color="auto"/></q:rPr></q:r>"#.to_vec());
    assert_eq!(
        inherited.underline_formatting().unwrap(),
        Some(RunUnderline {
            style: UnderlineStyle::Words,
            color: Some(RunUnderlineColor::Auto),
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
        })
    );
}

#[test]
fn rejects_invalid_or_duplicate_underline_properties() {
    for xml in [
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:val="triple"/></w:rPr></w:r>"#.as_slice(),
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:color="12345"/></w:rPr></w:r>"#.as_slice(),
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:themeColor="accent7"/></w:rPr></w:r>"#.as_slice(),
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u w:themeTint="000"/></w:rPr></w:r>"#.as_slice(),
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:u/><w:u/></w:rPr></w:r>"#.as_slice(),
    ] {
        assert!(Run::new(xml.to_vec()).underline_formatting().is_err());
    }
}

#[test]
fn rejects_unknown_entities_in_word_text() {
    let run = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>&unknown;</w:t></w:r>"#
            .to_vec(),
    );
    assert!(run.text().is_err());
    assert!(run.get_text_and_properties().is_err());
}

#[test]
fn omml_formulas_preserve_inline_and_display_xml_exactly() {
    let xml = br#"<wp:p xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:false="urn:not-omml">
            <wp:r>
                <m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>
                <q:oMath q:id="2"/>
            </wp:r>
            <m:oMathPara><q:oMath><q:r/></q:oMath></m:oMathPara>
            <false:oMath>ignored</false:oMath>
        </wp:p>"#;
    let paragraph = Paragraph::new(xml.to_vec());

    assert_eq!(
        paragraph.omml_formulas().unwrap(),
        vec![
            r#"<m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>"#,
            r#"<q:oMath q:id="2"/>"#,
        ]
    );
    assert_eq!(
        paragraph.paragraph_level_formulas().unwrap(),
        vec!["<q:oMath><q:r/></q:oMath>"]
    );
    assert_eq!(
        paragraph.runs().unwrap()[0].omml_formula().unwrap(),
        Some(
            r#"<m:oMath data-id="1"><m:r><wp:rPr/><m:t><![CDATA[x < y]]></m:t></m:r></m:oMath>"#
                .to_string()
        )
    );
    let inline = paragraph.inline_office_math().unwrap();
    assert_eq!(inline.len(), 2);
    assert_eq!(
        inline[1].xml(),
        r#"<q:oMath q:id="2" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"/>"#
    );
    assert_eq!(
        paragraph.display_office_math().unwrap()[0].xml(),
        r#"<q:oMath xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"><q:r/></q:oMath>"#
    );
}

#[test]
fn omml_formulas_accept_strict_and_inherited_prefixes() {
    let strict = Paragraph::new(
        br#"<s:p xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:math="http://purl.oclc.org/ooxml/officeDocument/math"><s:r><math:oMath><math:r/></math:oMath></s:r></s:p>"#
            .to_vec(),
    );
    assert_eq!(
        strict.omml_formulas().unwrap(),
        vec!["<math:oMath><math:r/></math:oMath>"]
    );

    let inherited = Run::new(br#"<w:r><m:oMath><m:r/></m:oMath></w:r>"#.to_vec());
    assert_eq!(
        inherited.omml_formula().unwrap().as_deref(),
        Some("<m:oMath><m:r/></m:oMath>")
    );

    let inherited_default = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:r><oMath><r/></oMath></w:r></w:p>"#
            .to_vec(),
    );
    assert_eq!(
        inherited_default.inline_office_math().unwrap()[0].xml(),
        r#"<oMath xmlns="http://schemas.openxmlformats.org/officeDocument/2006/math"><r/></oMath>"#
    );

    let foreign =
        Run::new(br#"<w:r xmlns:m="urn:not-omml"><m:oMath><m:r/></m:oMath></w:r>"#.to_vec());
    assert_eq!(foreign.omml_formula().unwrap(), None);
}

#[test]
fn omml_formulas_reject_malformed_xml() {
    let paragraph = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:r><m:oMath><m:r/></w:r></w:p>"#
            .to_vec(),
    );
    assert!(paragraph.omml_formulas().is_err());
    assert!(paragraph.paragraph_level_formulas().is_err());
}

#[test]
fn test_run_bold() {
    let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:b/></w:rPr>
            <w:t>Bold text</w:t>
        </w:r>"#;

    let run = Run::new(xml.to_vec());
    assert!(run.bold().unwrap().unwrap_or(false));
}

#[test]
fn test_run_italic() {
    let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:rPr><w:i/></w:rPr>
            <w:t>Italic text</w:t>
        </w:r>"#;

    let run = Run::new(xml.to_vec());
    assert!(run.italic().unwrap().unwrap_or(false));
}

#[test]
fn parses_typed_run_breaks_and_rendered_hints() {
    let xml = br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:t>Before</w:t><w:br/><w:br w:type="page"/><w:br w:type="column" w:clear="all"/>
            <w:br w:type="textWrapping" w:clear="left"></w:br>
            <w:lastRenderedPageBreak/><w:lastRenderedPageBreak></w:lastRenderedPageBreak>
        </w:r>"#;
    let run = Run::new(xml.to_vec());
    assert_eq!(
        run.breaks().unwrap().as_slice(),
        [
            RunBreak::default(),
            RunBreak {
                break_type: RunBreakType::Page,
                clear: RunBreakClear::None,
            },
            RunBreak {
                break_type: RunBreakType::Column,
                clear: RunBreakClear::All,
            },
            RunBreak {
                break_type: RunBreakType::TextWrapping,
                clear: RunBreakClear::Left,
            },
        ]
    );
    assert_eq!(run.last_rendered_page_break_count().unwrap(), 2);
    assert_eq!(run.text().unwrap(), "Before\n\n\n\n");
}

#[test]
fn rejects_invalid_run_break_enums() {
    let invalid_type = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:br w:type="section"/></w:r>"#
            .to_vec(),
    );
    assert!(invalid_type.breaks().is_err());

    let invalid_clear = Run::new(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:br w:clear="center"/></w:r>"#
            .to_vec(),
    );
    assert!(invalid_clear.breaks().is_err());
}

#[test]
fn reads_direct_paragraph_division_ids_namespace_aware() {
    let paragraph = Paragraph::new(
        br#"<q:p xmlns:q="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml"><q:pPr><false:divId false:val="9"/><q:divId q:val=" +123456789012345678901234567890 "/></q:pPr><q:r><q:t>text</q:t></q:r></q:p>"#
            .to_vec(),
    );
    assert_eq!(
        paragraph.division_id().unwrap().as_deref(),
        Some("+123456789012345678901234567890")
    );

    let inherited = Paragraph::new(br#"<q:p><q:pPr><q:divId q:val="-7"/></q:pPr></q:p>"#.to_vec());
    assert_eq!(inherited.division_id().unwrap().as_deref(), Some("-7"));

    let invalid = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:divId w:val="1.5"/></w:pPr></w:p>"#
            .to_vec(),
    );
    assert!(invalid.division_id().is_err());
}

#[test]
fn parses_typed_paragraph_spacing_attributes() {
    let paragraph = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="240" w:beforeLines="100" w:beforeAutospacing="off" w:after="120" w:afterLines="-25" w:afterAutospacing="on" w:line="-20" w:lineRule="exact"/></w:pPr><w:r><w:t>body</w:t></w:r></w:p>"#.to_vec(),
    );

    assert_eq!(
        paragraph.spacing().unwrap(),
        Some(ParagraphSpacing {
            before: Some(240),
            before_lines: Some(100),
            before_auto_spacing: Some(false),
            after: Some(120),
            after_lines: Some(-25),
            after_auto_spacing: Some(true),
            line: Some(-20),
            line_rule: Some(LineSpacingRule::Exact),
        })
    );
}

#[test]
fn rejects_invalid_paragraph_spacing_tokens_and_measurements() {
    for xml in [
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="-1"/></w:pPr></w:p>"#.as_slice(),
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:line="2147483648"/></w:pPr></w:p>"#.as_slice(),
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:lineRule="exactly"/></w:pPr></w:p>"#.as_slice(),
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:beforeAutospacing="maybe"/></w:pPr></w:p>"#.as_slice(),
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing/><w:spacing/></w:pPr></w:p>"#.as_slice(),
    ] {
        assert!(Paragraph::new(xml.to_vec()).spacing().is_err());
    }
}

#[test]
fn edits_typed_spacing_preserving_runs_and_paragraph_properties() {
    let mut paragraph = Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/></w:pPr><w:r><w:t>body</w:t></w:r></w:p>"#.to_vec(),
    );

    paragraph
        .set_spacing(Some(ParagraphSpacing {
            before: Some(240),
            line: Some(276),
            line_rule: Some(LineSpacingRule::Auto),
            ..ParagraphSpacing::default()
        }))
        .unwrap();
    assert_eq!(paragraph.spacing().unwrap().unwrap().before, Some(240));
    assert_eq!(paragraph.spacing().unwrap().unwrap().line, Some(276));
    assert_eq!(paragraph.style_id().unwrap().as_deref(), Some("Heading1"));
    assert_eq!(paragraph.text().unwrap(), "body");
    assert_eq!(paragraph.runs().unwrap().len(), 1);

    paragraph.set_spacing(None).unwrap();
    assert_eq!(paragraph.spacing().unwrap(), None);
    assert_eq!(paragraph.style_id().unwrap().as_deref(), Some("Heading1"));
    assert_eq!(paragraph.text().unwrap(), "body");
}

#[test]
fn edits_spacing_in_inherited_namespace_fragments_without_copying_runs() {
    let mut paragraph = Paragraph::new(br#"<q:p><q:r><q:t>body</q:t></q:r></q:p>"#.to_vec());

    paragraph
        .set_spacing(Some(ParagraphSpacing {
            after: Some(120),
            ..ParagraphSpacing::default()
        }))
        .unwrap();
    assert_eq!(paragraph.spacing().unwrap().unwrap().after, Some(120));
    assert_eq!(paragraph.text().unwrap(), "body");

    paragraph.set_spacing(None).unwrap();
    assert_eq!(paragraph.spacing().unwrap(), None);
    assert_eq!(paragraph.text().unwrap(), "body");
}
