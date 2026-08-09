#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    MathElementRole, MathObject, MathPropertiesKind, MathPropertyName, MathStructureChild,
    MathStructureKind, MathZoneKind, RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_inline_fraction_with_properties_and_round_trips() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi before{\mmath{\mf{\mfPr{\mtype bar}}{\mnum{\mr 1}}{\mden{\mr 2}}}} after}",
    )
    .unwrap();
    // Math run text stays in the typed tree and out of the body story.
    assert_eq!(document.text(), "before after");
    let zones = document.math_zones();
    assert_eq!(zones.len(), 1);
    let zone = &zones[0];
    assert_eq!(zone.kind, MathZoneKind::Inline);
    assert_eq!(zone.position, 6);
    assert_eq!(zone.content.len(), 1);
    let MathObject::Structure(fraction) = &zone.content[0] else {
        panic!("expected a fraction structure");
    };
    assert_eq!(fraction.kind, MathStructureKind::Fraction);
    let properties = fraction.properties.as_ref().unwrap();
    assert_eq!(
        properties.kind,
        MathPropertiesKind::Structure(MathStructureKind::Fraction)
    );
    assert_eq!(properties.properties.len(), 1);
    assert_eq!(properties.properties[0].name, MathPropertyName::Type);
    assert_eq!(properties.properties[0].value, "bar");
    assert_eq!(fraction.children.len(), 2);
    let MathStructureChild::Element(numerator) = &fraction.children[0] else {
        panic!("expected numerator element");
    };
    assert_eq!(numerator.role, MathElementRole::Numerator);
    let MathObject::Run(run) = &numerator.content[0] else {
        panic!("expected numerator run");
    };
    assert_eq!(run.text, "1");
    assert!(!run.normal_text);

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.math_zones(), zones);
}

#[test]
fn parses_display_zone_with_paragraph_properties() {
    let document =
        RtfDocument::parse(r"{\rtf1{\mmathPara{\mmathParaPr{\mjc centerGroup}}{\mr x}}}").unwrap();
    let zones = document.math_zones();
    assert_eq!(zones.len(), 1);
    assert_eq!(zones[0].kind, MathZoneKind::Display);
    let paragraph = zones[0].paragraph_properties.as_ref().unwrap();
    assert_eq!(paragraph.kind, MathPropertiesKind::Paragraph);
    assert_eq!(paragraph.properties[0].name, MathPropertyName::Justify);
    assert_eq!(paragraph.properties[0].value, "centerGroup");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_zones(), zones);
}

#[test]
fn parses_nested_structures_matrix_and_unicode() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\mmath"#,
        r#"{\mnary{\mnaryPr{\mchr \u8721?}{\msubHide 1}}{\msub }{\msup{\mr n}}{\me{\mr i}}}"#,
        r#"{\mrad{\mradPr{\mdegHide 1}}{\mdeg{\mr 3}}{\me{\mr x}}}"#,
        r#"{\mm{\mmPr{\mcount 2}{\mmcJc center}}{\mmr{\me{\mr a}}{\me{\mr b}}}{\mmr{\me{\mr c}}{\me{\mr d}}}}"#,
        r#"{\macc{\maccPr{\mchr \u770?}}{\me{\mr \mnor hat}}}"#,
        r#"{\msSubSup{\msub{\mr 1}}{\msup{\mr 2}}{\me{\mr x}}}"#,
        r"}}",
    ))
    .unwrap();
    assert_eq!(document.text(), "");
    let zones = document.math_zones();
    assert_eq!(zones.len(), 1);
    assert_eq!(zones[0].content.len(), 5);

    let MathObject::Structure(nary) = &zones[0].content[0] else {
        panic!("expected n-ary structure");
    };
    assert_eq!(nary.kind, MathStructureKind::Nary);
    let nary_pr = nary.properties.as_ref().unwrap();
    assert_eq!(nary_pr.properties[0].name, MathPropertyName::Char);
    assert_eq!(nary_pr.properties[0].value, "\u{2211}");
    assert_eq!(nary_pr.properties[1].name, MathPropertyName::SubscriptHide);
    assert_eq!(nary_pr.properties[1].value, "1");

    let MathObject::Structure(matrix) = &zones[0].content[2] else {
        panic!("expected matrix structure");
    };
    assert_eq!(matrix.kind, MathStructureKind::Matrix);
    assert_eq!(matrix.children.len(), 2);
    let MathStructureChild::MatrixRow(row) = &matrix.children[0] else {
        panic!("expected matrix row");
    };
    assert_eq!(row.cells.len(), 2);

    let MathObject::Structure(accent) = &zones[0].content[3] else {
        panic!("expected accent structure");
    };
    let MathStructureChild::Element(accent_element) = &accent.children[0] else {
        panic!("expected accent element");
    };
    let MathObject::Run(accent_run) = &accent_element.content[0] else {
        panic!("expected accent run");
    };
    assert!(accent_run.normal_text);
    assert_eq!(accent_run.text, "hat");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_zones(), zones);
}

#[test]
fn skips_mmath_pict_fallback_and_accepts_momath_aliases() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\moMath{\mr 1}{\*\mmathPict {\pict\pngblip 00}}}"#,
        r#"{\moMathPara{\moMathParaPr{\mjc center}}{\mr 2}}}"#,
    ))
    .unwrap();
    let zones = document.math_zones();
    assert_eq!(zones.len(), 2);
    assert_eq!(zones[0].kind, MathZoneKind::Inline);
    assert_eq!(zones[1].kind, MathZoneKind::Display);
    let paragraph = zones[1].paragraph_properties.as_ref().unwrap();
    assert_eq!(paragraph.properties[0].name, MathPropertyName::Justify);

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_zones(), zones);
}

#[test]
fn coexists_with_body_markup_positions() {
    let document =
        RtfDocument::parse(r"{\rtf1 ab{\mmath{\mr 1}}cd{\*\bkmkstart bm}ef{\*\bkmkend bm}}")
            .unwrap();
    assert_eq!(document.text(), "abcdef");
    assert_eq!(document.math_zones()[0].position, 2);
    let bookmark = &document.bookmarks().bookmarks()[0];
    assert_eq!(bookmark.position, 4);
    assert_eq!(bookmark.content, "ef");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_zones(), document.math_zones());
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_constructors_validate_structure() {
    use litchi_rtf::{MathProperties, MathStructure, MathZone};
    use std::borrow::Cow;
    // Inline zones cannot carry paragraph properties.
    assert!(
        MathZone::new(
            MathZoneKind::Inline,
            Some(MathProperties::new(MathPropertiesKind::Paragraph, Vec::new()).unwrap()),
            Vec::new(),
            0,
        )
        .is_err()
    );
    // Display zones must use the mmathParaPr destination.
    assert!(
        MathZone::new(
            MathZoneKind::Display,
            Some(MathProperties::new(MathPropertiesKind::Run, Vec::new()).unwrap()),
            Vec::new(),
            0,
        )
        .is_err()
    );
    // A fraction requires numerator and denominator children.
    assert!(MathStructure::new(MathStructureKind::Fraction, None, Vec::new()).is_err());
    // Property values reject control characters.
    assert!(litchi_rtf::MathProperty::new(MathPropertyName::Type, Cow::Borrowed("ba\rr")).is_err());
}

#[test]
fn rejects_misplaced_or_malformed_math() {
    let cases = [
        // Structure destination outside a math zone.
        r"{\rtf1{\mf{\mnum{\mr 1}}{\mden{\mr 2}}}}",
        // Bare math control in the body.
        r"{\rtf1\mf x}",
        // Missing required denominator.
        r"{\rtf1{\mmath{\mf{\mnum{\mr 1}}}}}",
        // Wrong argument role for a fraction.
        r"{\rtf1{\mmath{\mf{\mdeg{\mr 1}}{\mden{\mr 2}}}}}",
        // Duplicate property names in one destination.
        r"{\rtf1{\mmath{\mf{\mfPr{\mtype bar}{\mtype f}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Unsupported property control.
        r"{\rtf1{\mmath{\mf{\mfPr{\mfoo bar}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Ungrouped text inside a zone.
        r"{\rtf1{\mmath x}}",
        // Matrix row outside a matrix.
        r"{\rtf1{\mmath{\mf{\mmr{\me{\mr 1}}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Property with conflicting parameter and text values.
        r"{\rtf1{\mmath{\mf{\mfPr{\mgrow1 x}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Property destination after the children.
        r"{\rtf1{\mmath{\mf{\mnum{\mr 1}}{\mden{\mr 2}}{\mfPr{\mtype bar}}}}}",
        // Paragraph properties in an inline zone.
        r"{\rtf1{\mmath{\mmathParaPr{\mjc center}}{\mr x}}}",
        // Binary data inside a math run.
        r"{\rtf1{\mmath{\mr \bin2 xx}}}",
        // Unterminated zone.
        r"{\rtf1{\mmath{\mf{\mnum{\mr 1}}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_matrix_columns_and_argument_properties() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\mmath"#,
        r#"{\mm{\mmPr{\mcount 2}{\mmcs{\mmc{\mmcPr{\mcount 2}{\mmcJc center}}}{\mmc{\mmcPr{\mcount 1}}}}}"#,
        r#"{\mmr{\me{\mr a}}{\me{\mr b}}}}"#,
        r#"{\msSup{\msup{\mr 2}}{\me{\margPr{\margSz 2}}{\mr x}}}"#,
        r"}}",
    ))
    .unwrap();
    let zones = document.math_zones();
    assert_eq!(zones.len(), 1);

    let MathObject::Structure(matrix) = &zones[0].content[0] else {
        panic!("expected matrix structure");
    };
    let matrix_pr = matrix.properties.as_ref().unwrap();
    assert_eq!(matrix_pr.matrix_columns.len(), 2);
    let first_column = matrix_pr.matrix_columns[0].properties.as_ref().unwrap();
    assert_eq!(first_column.kind, MathPropertiesKind::MatrixColumn);
    assert_eq!(
        first_column.properties[0].name,
        MathPropertyName::MatrixCellCount
    );
    assert_eq!(first_column.properties[0].value, "2");
    assert_eq!(
        first_column.properties[1].name,
        MathPropertyName::MatrixCellJustify
    );
    assert_eq!(first_column.properties[1].value, "center");
    let second_column = matrix_pr.matrix_columns[1].properties.as_ref().unwrap();
    assert_eq!(second_column.properties.len(), 1);

    let MathObject::Structure(superscript) = &zones[0].content[1] else {
        panic!("expected superscript structure");
    };
    let MathStructureChild::Element(base) = &superscript.children[1] else {
        panic!("expected base element");
    };
    let argument = base.argument_properties.as_ref().unwrap();
    assert_eq!(argument.kind, MathPropertiesKind::Argument);
    assert_eq!(argument.properties[0].name, MathPropertyName::ArgumentSize);
    assert_eq!(argument.properties[0].value, "2");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.math_zones(), zones);
}

#[test]
fn rejects_misplaced_matrix_columns_and_argument_properties() {
    let cases = [
        // Argument properties after argument content.
        r"{\rtf1{\mmath{\msSup{\msup{\mr 2}}{\me{\mr x}{\margPr{\margSz 1}}}}}}",
        // Argument size outside an argument-properties destination.
        r"{\rtf1{\mmath{\mf{\mfPr{\margSz 1}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Non-argument property inside margPr.
        r"{\rtf1{\mmath{\msSup{\msup{\mr 2}}{\me{\margPr{\mtype bar}}{\mr x}}}}}",
        // Matrix columns outside matrix properties.
        r"{\rtf1{\mmath{\mf{\mfPr{\mmcs{\mmc}}}{\mnum{\mr 1}}{\mden{\mr 2}}}}}",
        // Empty matrix columns destination.
        r"{\rtf1{\mmath{\mm{\mmPr{\mmcs }}{\mmr{\me{\mr a}}}}}}",
        // Matrix column with an unsupported group.
        r"{\rtf1{\mmath{\mm{\mmPr{\mmcs{\mmc{\mf{\mnum{\mr 1}}{\mden{\mr 2}}}}}}{\mmr{\me{\mr a}}}}}}",
        // Non-column property inside mmcPr.
        r"{\rtf1{\mmath{\mm{\mmPr{\mmcs{\mmc{\mmcPr{\mtype bar}}}}}{\mmr{\me{\mr a}}}}}}",
        // Duplicate matrix columns destinations.
        r"{\rtf1{\mmath{\mm{\mmPr{\mmcs{\mmc}}{\mmcs{\mmc}}}{\mmr{\me{\mr a}}}}}}",
        // Matrix column destinations at zone level.
        r"{\rtf1{\mmath{\mmcs{\mmc}}}}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn typed_constructors_validate_new_property_scopes() {
    use litchi_rtf::{MathProperties, MathProperty};
    use std::borrow::Cow;
    let argument_size =
        MathProperty::new(MathPropertyName::ArgumentSize, Cow::Borrowed("2")).unwrap();
    // \margSz is only permitted inside \margPr.
    assert!(
        MathProperties::new(
            MathPropertiesKind::Structure(MathStructureKind::Fraction),
            vec![argument_size.clone()],
        )
        .is_err()
    );
    // \margPr accepts nothing but \margSz.
    let fraction_type = MathProperty::new(MathPropertyName::Type, Cow::Borrowed("bar")).unwrap();
    assert!(MathProperties::new(MathPropertiesKind::Argument, vec![fraction_type]).is_err());
    assert!(MathProperties::new(MathPropertiesKind::Argument, vec![argument_size]).is_ok());
}

#[test]
fn rejects_excessive_math_nesting_depth() {
    let mut rtf = String::from(r"{\rtf1{\mmath");
    for _ in 0..65 {
        rtf.push_str(r"{\mbox{\me");
    }
    rtf.push_str(r"{\mr 1}");
    for _ in 0..65 {
        rtf.push_str("}}");
    }
    rtf.push_str("}}");
    assert!(RtfDocument::parse(&rtf).is_err());

    let mut valid = String::from(r"{\rtf1{\mmath");
    for _ in 0..62 {
        valid.push_str(r"{\mbox{\me");
    }
    valid.push_str(r"{\mr 1}");
    for _ in 0..62 {
        valid.push_str("}}");
    }
    valid.push_str("}}");
    let document = RtfDocument::parse(&valid).unwrap();
    assert_eq!(document.math_zones().len(), 1);
}
