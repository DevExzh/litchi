use litchi_rtf::{
    DocumentMathProperties, MathBinaryOperatorBreak, MathBinarySubtractionBreak, MathFlag,
    MathJustification, MathLimitPlacement, RtfDocument, RtfWriter,
};
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_libreoffice_defaults_as_inert_metadata_and_round_trips() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\mmathPr\mmathFont34\mbrkBin0\mbrkBinSub0"#,
        r#"\msmallFrac0\mdispDef1\mlMargin0\mrMargin0\mdefJc1"#,
        r#"\mwrapIndent1440\mintLim0\mnaryLim1}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let properties = document.math_properties().unwrap();
    assert_eq!(properties.math_font, Some(34));
    assert_eq!(
        properties.binary_operator_break,
        Some(MathBinaryOperatorBreak::Before)
    );
    assert_eq!(
        properties.default_justification,
        Some(MathJustification::CenteredAsGroup)
    );
    assert_eq!(properties.wrap_indent, Some(1440));
    assert_eq!(
        properties.nary_limit_placement,
        Some(MathLimitPlacement::UnderOver)
    );

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_properties(), Some(properties));
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn accepts_unstarred_math_properties_destination() {
    let document =
        RtfDocument::parse(r#"{\rtf1{\mmathPr\mmathFont34\mdefJc1\mwrapIndent1440}Body}"#).unwrap();
    let properties = document.math_properties().unwrap();
    assert_eq!(properties.math_font, Some(34));
    assert_eq!(
        properties.default_justification,
        Some(MathJustification::CenteredAsGroup)
    );
    assert_eq!(properties.wrap_indent, Some(1440));
    assert_eq!(document.text(), "Body");
}

#[test]
fn preserves_complete_group_and_unknown_finite_values() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\mmathPr\mbrkBin17\mbrkBinSub-3\mdefJc99\mdispDef8"#,
        r#"\minterSp20\mintLim7\mintraSp30\mlMargin40\mmathFont5"#,
        r#"\mnaryLim6\mpostSp50\mpreSp60\mrMargin70\msmallFrac4"#,
        r#"\mwrapIndent80\mwrapRight9}X}"#,
    ))
    .unwrap();
    let properties = document.math_properties().unwrap();
    assert_eq!(
        properties.binary_operator_break,
        Some(MathBinaryOperatorBreak::Unknown(17))
    );
    assert_eq!(
        properties.binary_subtraction_break,
        Some(MathBinarySubtractionBreak::Unknown(-3))
    );
    assert_eq!(
        properties.default_justification,
        Some(MathJustification::Unknown(99))
    );
    assert_eq!(properties.display_defaults, Some(MathFlag::Unknown(8)));
    assert_eq!(
        properties.integral_limit_placement,
        Some(MathLimitPlacement::Unknown(7))
    );
    assert_eq!(properties.inter_equation_spacing, Some(20));
    assert_eq!(properties.intra_equation_spacing, Some(30));
    assert_eq!(properties.post_spacing, Some(50));
    assert_eq!(properties.pre_spacing, Some(60));
    assert_eq!(properties.wrap_right, Some(MathFlag::Unknown(9)));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"\mbrkBin17"#));
    assert!(serialized.contains(r#"\mbrkBinSub-3"#));
    assert!(serialized.contains(r#"\mwrapRight9"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.math_properties(), Some(properties));
}

#[test]
fn mutation_and_clear_preserve_body() {
    let mut properties = DocumentMathProperties::new();
    properties.math_font = Some(7);
    properties.binary_operator_break = Some(MathBinaryOperatorBreak::After);
    properties.binary_subtraction_break = Some(MathBinarySubtractionBreak::PlusMinus);
    properties.default_justification = Some(MathJustification::Right);
    properties.display_defaults = Some(MathFlag::On);
    properties.integral_limit_placement = Some(MathLimitPlacement::UnderOver);
    properties.small_fractions = Some(MathFlag::Off);
    properties.wrap_right = Some(MathFlag::On);

    let mut document = RtfDocument::parse(r#"{\rtf1 Text}"#).unwrap();
    document.set_math_properties(properties.clone()).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_properties(), Some(&properties));
    assert_eq!(reparsed.text(), "Text");

    document.clear_math_properties();
    assert!(document.math_properties().is_none());
    assert_eq!(document.text(), "Text");

    let mut invalid = DocumentMathProperties::new();
    invalid.math_font = Some(i32::MAX as u32 + 1);
    assert!(document.set_math_properties(invalid).is_err());
}

#[test]
fn rejects_malformed_or_active_math_properties() {
    let cases = [
        r#"{\rtf1{\*\mmathPr}{\*\mmathPr}}"#,
        r#"{\rtf1{\*\mmathPr\mbrkBin0\mbrkBin1}}"#,
        r#"{\rtf1{\*\mmathPr text}}"#,
        r#"{\rtf1{\*\mmathPr{\mbrkBin0}}}"#,
        r#"{\rtf1{\*\mmathPr\bin2 xx}}"#,
        r#"{\rtf1{\*\mmathPr\b}}"#,
        r#"{\rtf1\mbrkBin0}"#,
        r#"{\rtf1{\*\mmathPr\mmathFont-1}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_math_property_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/rtfexport/data/tdf161878.rtf",
        "sw/qa/extras/odfexport/data/footnote_spacing_hanging_para.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let properties = document
            .math_properties()
            .unwrap_or_else(|| panic!("fixture exposed no math properties: {fixture}"));
        assert_eq!(properties.math_font, Some(34));
        assert_eq!(properties.wrap_indent, Some(1440));
    }
}
