// OMML writer tests: per-construct serialization plus parse/write round-trips

use super::OmmlWriter;
use crate::ast::{
    AccentType, Fence, Formula, FractionType, LargeOperator, MathNode, MatrixFence, Operator,
    PredefinedSymbol, StyleType,
};
use crate::omml::OmmlParser;
use std::borrow::Cow;

/// Serialize AST nodes and return the XML
fn write(nodes: &[MathNode]) -> String {
    let mut writer = OmmlWriter::new();
    writer
        .write_nodes(nodes)
        .expect("serialization should succeed")
        .to_string()
}

/// Parse OMML, re-serialize it, re-parse, and assert the ASTs match
fn assert_roundtrip(xml: &str) {
    let formula1 = Formula::new();
    let parser1 = OmmlParser::new(formula1.arena());
    let ast1 = parser1.parse(xml).expect("initial parse should succeed");

    let mut writer = OmmlWriter::new();
    let serialized = writer
        .write_nodes(&ast1)
        .expect("serialization should succeed")
        .to_string();

    let formula2 = Formula::new();
    let parser2 = OmmlParser::new(formula2.arena());
    let ast2 = parser2
        .parse(&serialized)
        .unwrap_or_else(|e| panic!("re-parse failed: {e}\nserialized: {serialized}"));

    assert_eq!(ast1, ast2, "round-trip mismatch\nserialized: {serialized}");
}

fn text(value: &'static str) -> MathNode<'static> {
    MathNode::Text(Cow::Borrowed(value))
}

// ---------------------------------------------------------------------------
// Unit tests: emitted XML per construct
// ---------------------------------------------------------------------------

#[test]
fn writes_text_run() {
    let xml = write(&[text("x")]);
    assert_eq!(
        xml,
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">\
         <m:r><m:t>x</m:t></m:r></m:oMath>"
    );
}

#[test]
fn writes_operator_and_symbol_runs() {
    let xml = write(&[
        MathNode::Operator(Operator::PlusMinus),
        MathNode::PredefinedSymbol(PredefinedSymbol::Alpha),
        MathNode::Number(Cow::Borrowed("42")),
    ]);
    assert!(xml.contains("<m:r><m:t>±</m:t></m:r>"));
    assert!(xml.contains("<m:r><m:t>α</m:t></m:r>"));
    assert!(xml.contains("<m:r><m:t>42</m:t></m:r>"));
}

#[test]
fn writes_fraction_with_type() {
    let xml = write(&[MathNode::Frac {
        numerator: vec![text("1")],
        denominator: vec![text("2")],
        line_thickness: None,
        frac_type: Some(FractionType::NoBar),
    }]);
    assert!(xml.contains("<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr>"));
    assert!(xml.contains("<m:num><m:r><m:t>1</m:t></m:r></m:num>"));
    assert!(xml.contains("<m:den><m:r><m:t>2</m:t></m:r></m:den>"));
}

#[test]
fn writes_radical_with_hidden_degree() {
    let xml = write(&[MathNode::Root {
        base: vec![text("x")],
        index: None,
    }]);
    assert!(xml.contains("<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>"));
    assert!(xml.contains("<m:e><m:r><m:t>x</m:t></m:r></m:e>"));
}

#[test]
fn writes_radical_with_degree() {
    let xml = write(&[MathNode::Root {
        base: vec![text("x")],
        index: Some(vec![text("3")]),
    }]);
    assert!(xml.contains("<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg>"));
    assert!(!xml.contains("degHide"));
}

#[test]
fn writes_scripts() {
    let power = write(&[MathNode::Power {
        base: vec![text("x")],
        exponent: vec![text("2")],
    }]);
    assert!(power.contains("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup>"));

    let sub = write(&[MathNode::Sub {
        base: vec![text("a")],
        subscript: vec![text("i")],
    }]);
    assert!(sub.contains("<m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub>"));

    let subsup = write(&[MathNode::SubSup {
        base: vec![text("x")],
        subscript: vec![text("i")],
        superscript: vec![text("2")],
    }]);
    assert!(subsup.contains("<m:sSubSup>"));
    assert!(subsup.contains("<m:sub><m:r><m:t>i</m:t></m:r></m:sub>"));
    assert!(subsup.contains("<m:sup><m:r><m:t>2</m:t></m:r></m:sup>"));
}

#[test]
fn writes_pre_scripts() {
    let xml = write(&[MathNode::PreSubSup {
        base: vec![text("X")],
        pre_subscript: vec![text("n")],
        pre_superscript: vec![text("m")],
    }]);
    assert!(xml.contains(
        "<m:sPre><m:sub><m:r><m:t>n</m:t></m:r></m:sub>\
         <m:sup><m:r><m:t>m</m:t></m:r></m:sup><m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre>"
    ));

    let pre_sub = write(&[MathNode::PreSub {
        base: vec![text("X")],
        pre_subscript: vec![text("n")],
    }]);
    assert!(pre_sub.contains("<m:sub><m:r><m:t>n</m:t></m:r></m:sub><m:sup/>"));
}

#[test]
fn writes_delimiter_with_fences() {
    let xml = write(&[MathNode::Fenced {
        open: Fence::Bracket,
        content: vec![text("x")],
        close: Fence::Bracket,
        separator: None,
    }]);
    assert!(xml.contains(
        "<m:d><m:dPr><m:begChr m:val=\"[\"/><m:endChr m:val=\"]\"/></m:dPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>"
    ));
}

#[test]
fn writes_box_for_fenceless_group() {
    let xml = write(&[MathNode::Fenced {
        open: Fence::None,
        content: vec![text("x")],
        close: Fence::None,
        separator: None,
    }]);
    assert!(xml.contains("<m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box>"));
}

#[test]
fn writes_nary_with_hidden_limits() {
    let xml = write(&[MathNode::LargeOp {
        operator: LargeOperator::Integral,
        lower_limit: None,
        upper_limit: None,
        integrand: Some(vec![text("x")]),
        hide_lower: true,
        hide_upper: true,
    }]);
    assert!(xml.contains(
        "<m:naryPr><m:chr m:val=\"∫\"/><m:subHide m:val=\"1\"/>\
         <m:supHide m:val=\"1\"/></m:naryPr><m:sub/><m:sup/>"
    ));
}

#[test]
fn writes_function() {
    let xml = write(&[MathNode::Function {
        name: Cow::Borrowed("sin"),
        argument: vec![text("x")],
    }]);
    assert!(xml.contains(
        "<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>"
    ));
}

#[test]
fn writes_matrix_rows_and_cells() {
    let xml = write(&[MathNode::Matrix {
        rows: vec![
            vec![vec![text("a")], vec![text("b")]],
            vec![vec![text("c")], vec![text("d")]],
        ],
        fence_type: MatrixFence::None,
        properties: None,
    }]);
    assert!(xml.contains(
        "<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e>\
         <m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
    ));
}

#[test]
fn writes_fenced_matrix_as_delimited() {
    let xml = write(&[MathNode::Matrix {
        rows: vec![vec![vec![text("a")]]],
        fence_type: MatrixFence::Paren,
        properties: None,
    }]);
    assert!(xml.contains("<m:d><m:dPr><m:begChr m:val=\"(\"/><m:endChr m:val=\")\"/></m:dPr>"));
    assert!(xml.contains("<m:e><m:m><m:mr>"));
}

#[test]
fn writes_eq_array_with_properties() {
    use crate::ast::{Alignment, EqArrayProperties};
    let xml = write(&[MathNode::EqArray {
        rows: vec![vec![text("a")], vec![text("b")]],
        properties: Some(EqArrayProperties {
            base_alignment: Some(Alignment::Center),
            max_distance: None,
            object_distance: None,
            row_spacing: None,
            row_spacing_rule: None,
        }),
    }]);
    assert!(xml.contains("<m:eqArr><m:eqArrPr><m:baseJc m:val=\"center\"/></m:eqArrPr>"));
    assert!(xml.contains("<m:e><m:r><m:t>a</m:t></m:r></m:e>"));
}

#[test]
fn writes_accent() {
    let xml = write(&[MathNode::Accent {
        base: Box::new(vec![text("x")]),
        accent: AccentType::Hat,
        position: None,
    }]);
    assert!(xml.contains("<m:acc><m:accPr><m:chr m:val=\"\u{0302}\"/></m:accPr>"));
}

#[test]
fn writes_bar_with_position() {
    use crate::ast::Position;
    let xml = write(&[MathNode::Bar {
        base: Box::new(vec![text("x")]),
        position: Some(Position::Bottom),
    }]);
    assert!(xml.contains("<m:bar><m:barPr><m:pos m:val=\"bot\"/></m:barPr>"));
}

#[test]
fn writes_group_char() {
    use crate::ast::{Position, VerticalAlignment};
    let xml = write(&[MathNode::GroupChar {
        base: Box::new(vec![text("x")]),
        character: Some(Cow::Borrowed("⏟")),
        position: Some(Position::Bottom),
        vertical_alignment: Some(VerticalAlignment::Top),
    }]);
    assert!(xml.contains(
        "<m:groupChr><m:groupChrPr><m:chr m:val=\"⏟\"/>\
         <m:pos m:val=\"bot\"/><m:vertJc m:val=\"top\"/></m:groupChrPr>"
    ));
}

#[test]
fn writes_under_over() {
    let under = write(&[MathNode::Under {
        base: vec![text("lim")],
        under: vec![text("n")],
        position: None,
    }]);
    assert!(under.contains(
        "<m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e>\
         <m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow>"
    ));

    let over = write(&[MathNode::Over {
        base: vec![text("x")],
        over: vec![text("~")],
        position: None,
    }]);
    assert!(over.contains("<m:limUpp>"));
}

#[test]
fn writes_under_over_combined() {
    let xml = write(&[MathNode::UnderOver {
        base: vec![text("x")],
        under: vec![text("a")],
        over: vec![text("b")],
        position: None,
    }]);
    // Nested limUpp (over) inside limLow (under)
    assert!(xml.contains("<m:limLow><m:e><m:limUpp>"));
    assert!(xml.contains("<m:lim><m:r><m:t>a</m:t></m:r></m:lim></m:limLow>"));
}

#[test]
fn writes_run_with_properties() {
    let xml = write(&[MathNode::Run {
        content: vec![text("abc")],
        literal: Some(true),
        style: Some(StyleType::Bold),
        font: Some(Cow::Borrowed("Cambria Math")),
        color: None,
        underline: None,
        overline: None,
        strike_through: None,
        double_strike_through: None,
    }]);
    assert!(xml.contains(
        "<m:r><m:rPr><m:lit m:val=\"1\"/><m:sty m:val=\"b\"/>\
         <m:nor>Cambria Math</m:nor></m:rPr><m:t>abc</m:t></m:r>"
    ));
}

#[test]
fn writes_border_box_style_flags() {
    use crate::ast::BorderBoxStyle;
    let xml = write(&[MathNode::BorderBox {
        content: Box::new(vec![text("x")]),
        style: Some(BorderBoxStyle {
            hide_top: true,
            hide_bottom: false,
            hide_left: false,
            hide_right: false,
            strike_horizontal: true,
            strike_vertical: false,
            strike_bltr: false,
            strike_tlbr: false,
        }),
    }]);
    assert!(xml.contains(
        "<m:borderBox><m:borderBoxPr><m:hideTop m:val=\"1\"/>\
         <m:strikeH m:val=\"1\"/></m:borderBoxPr>"
    ));
}

#[test]
fn escapes_xml_special_characters() {
    let xml = write(&[text("a<b&c>d")]);
    assert!(xml.contains("<m:t>a&lt;b&amp;c&gt;d</m:t>"));
}

#[test]
fn error_node_fails_serialization() {
    let mut writer = OmmlWriter::new();
    let nodes = [MathNode::Error(Cow::Borrowed("bad input"))];
    let result = writer.write_nodes(&nodes);
    assert!(result.is_err());
}

// ---------------------------------------------------------------------------
// Round-trip tests: OMML -> AST -> OMML -> AST must preserve the AST
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_text() {
    assert_roundtrip("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
}

#[test]
fn roundtrip_multiple_runs() {
    assert_roundtrip(
        "<m:oMath><m:r><m:t>a</m:t></m:r><m:r><m:t>+</m:t></m:r>\
         <m:r><m:t>b</m:t></m:r></m:oMath>",
    );
}

#[test]
fn roundtrip_escaped_text() {
    assert_roundtrip("<m:oMath><m:r><m:t>a&lt;b&amp;c</m:t></m:r></m:oMath>");
}

#[test]
fn roundtrip_fraction() {
    assert_roundtrip(
        "<m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num>\
         <m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>",
    );
}

#[test]
fn roundtrip_fraction_nobar() {
    assert_roundtrip(
        "<m:oMath><m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr>\
         <m:num><m:r><m:t>a</m:t></m:r></m:num>\
         <m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>",
    );
}

#[test]
fn roundtrip_superscript() {
    assert_roundtrip(
        "<m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>\
         <m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>",
    );
}

#[test]
fn roundtrip_subscript() {
    assert_roundtrip(
        "<m:oMath><m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e>\
         <m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:oMath>",
    );
}

#[test]
fn roundtrip_subsup() {
    assert_roundtrip(
        "<m:oMath><m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e>\
         <m:sub><m:r><m:t>i</m:t></m:r></m:sub>\
         <m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup></m:oMath>",
    );
}

#[test]
fn roundtrip_radical() {
    assert_roundtrip(
        "<m:oMath><m:rad><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>",
    );
}

#[test]
fn roundtrip_radical_with_degree() {
    assert_roundtrip(
        "<m:oMath><m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>",
    );
}

#[test]
fn roundtrip_nary_sum_with_limits() {
    assert_roundtrip(
        "<m:oMath><m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr>\
         <m:sub><m:r><m:t>i</m:t></m:r></m:sub>\
         <m:sup><m:r><m:t>n</m:t></m:r></m:sup>\
         <m:e><m:r><m:t>a</m:t></m:r></m:e></m:nary></m:oMath>",
    );
}

#[test]
fn roundtrip_nary_hidden_limits() {
    assert_roundtrip(
        "<m:oMath><m:nary><m:naryPr><m:chr m:val=\"∫\"/>\
         <m:subHide m:val=\"1\"/><m:supHide m:val=\"1\"/></m:naryPr>\
         <m:sub/><m:sup/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>",
    );
}

#[test]
fn roundtrip_delimiter_brackets() {
    assert_roundtrip(
        "<m:oMath><m:d><m:dPr><m:begChr m:val=\"[\"/><m:endChr m:val=\"]\"/></m:dPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:oMath>",
    );
}

#[test]
fn roundtrip_delimiter_with_separator() {
    assert_roundtrip(
        "<m:oMath><m:d><m:dPr><m:begChr m:val=\"(\"/><m:sepChr m:val=\"|\"/>\
         <m:endChr m:val=\")\"/></m:dPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:oMath>",
    );
}

#[test]
fn roundtrip_norm_delimiter() {
    assert_roundtrip(
        "<m:oMath><m:d><m:dPr><m:begChr m:val=\"‖\"/><m:endChr m:val=\"‖\"/></m:dPr>\
         <m:e><m:r><m:t>v</m:t></m:r></m:e></m:d></m:oMath>",
    );
}

#[test]
fn roundtrip_function() {
    assert_roundtrip(
        "<m:oMath><m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>",
    );
}

#[test]
fn roundtrip_matrix_with_multi_node_cells() {
    assert_roundtrip(
        "<m:oMath><m:m>\
         <m:mr><m:e><m:r><m:t>a</m:t></m:r><m:r><m:t>+</m:t></m:r>\
         <m:r><m:t>b</m:t></m:r></m:e><m:e><m:r><m:t>c</m:t></m:r></m:e></m:mr>\
         <m:mr><m:e><m:r><m:t>d</m:t></m:r></m:e><m:e><m:r><m:t>e</m:t></m:r></m:e></m:mr>\
         </m:m></m:oMath>",
    );
}

#[test]
fn roundtrip_eq_array_with_base_jc() {
    assert_roundtrip(
        "<m:oMath><m:eqArr><m:eqArrPr><m:baseJc m:val=\"center\"/></m:eqArrPr>\
         <m:e><m:r><m:t>a</m:t></m:r></m:e>\
         <m:e><m:r><m:t>b</m:t></m:r></m:e></m:eqArr></m:oMath>",
    );
}

#[test]
fn roundtrip_accent_val_attribute_form() {
    // Word writes the accent character via an m:val attribute
    assert_roundtrip(
        "<m:oMath><m:acc><m:accPr><m:chr m:val=\"\u{0302}\"/></m:accPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>",
    );
}

#[test]
fn roundtrip_bar_bottom() {
    assert_roundtrip(
        "<m:oMath><m:bar><m:barPr><m:pos m:val=\"bot\"/></m:barPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar></m:oMath>",
    );
}

#[test]
fn roundtrip_group_char() {
    assert_roundtrip(
        "<m:oMath><m:groupChr><m:groupChrPr><m:chr m:val=\"⏟\"/>\
         <m:pos m:val=\"bot\"/><m:vertJc m:val=\"top\"/></m:groupChrPr>\
         <m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr></m:oMath>",
    );
}

#[test]
fn roundtrip_phantom() {
    assert_roundtrip(
        "<m:oMath><m:phant><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant></m:oMath>",
    );
}

#[test]
fn roundtrip_box() {
    assert_roundtrip(
        "<m:oMath><m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box></m:oMath>",
    );
}

#[test]
fn roundtrip_run_properties() {
    assert_roundtrip(
        "<m:oMath><m:r><m:rPr><m:lit m:val=\"1\"/><m:scr m:val=\"script\"/>\
         <m:nor>Cambria Math</m:nor></m:rPr><m:t>abc</m:t></m:r></m:oMath>",
    );
}

#[test]
fn roundtrip_run_style_bold_italic() {
    assert_roundtrip(
        "<m:oMath><m:r><m:rPr><m:sty m:val=\"bi\"/></m:rPr>\
         <m:t>x</m:t></m:r></m:oMath>",
    );
}

#[test]
fn roundtrip_pre_scripts() {
    assert_roundtrip(
        "<m:oMath><m:sPre><m:sub><m:r><m:t>n</m:t></m:r></m:sub>\
         <m:sup><m:r><m:t>m</m:t></m:r></m:sup>\
         <m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre></m:oMath>",
    );
}

#[test]
fn roundtrip_pre_subscript_only() {
    assert_roundtrip(
        "<m:oMath><m:sPre><m:sub><m:r><m:t>n</m:t></m:r></m:sub><m:sup/>\
         <m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre></m:oMath>",
    );
}

#[test]
fn roundtrip_lim_low() {
    assert_roundtrip(
        "<m:oMath><m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e>\
         <m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow></m:oMath>",
    );
}

#[test]
fn roundtrip_lim_upp() {
    assert_roundtrip(
        "<m:oMath><m:limUpp><m:e><m:r><m:t>x</m:t></m:r></m:e>\
         <m:lim><m:r><m:t>+</m:t></m:r></m:lim></m:limUpp></m:oMath>",
    );
}

#[test]
fn roundtrip_complex_nested_formula() {
    assert_roundtrip(
        "<m:oMath><m:sSup><m:e><m:func>\
         <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>\
         <m:e><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num>\
         <m:den><m:rad><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad></m:den></m:f></m:e>\
         </m:func></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>",
    );
}

// ---------------------------------------------------------------------------
// LaTeX consistency: manually built ASTs survive OMML round-trips with the
// same LaTeX rendering
// ---------------------------------------------------------------------------

fn latex_of(nodes: &[MathNode]) -> String {
    let mut converter = crate::latex::LatexConverter::new();
    converter
        .convert_nodes(nodes)
        .expect("LaTeX conversion should succeed")
        .to_string()
}

fn assert_latex_consistent(nodes: &[MathNode]) {
    let latex_before = latex_of(nodes);

    let mut writer = OmmlWriter::new();
    let serialized = writer
        .write_nodes(nodes)
        .expect("serialization should succeed")
        .to_string();

    let formula = Formula::new();
    let parser = OmmlParser::new(formula.arena());
    let reparsed = parser
        .parse(&serialized)
        .unwrap_or_else(|e| panic!("re-parse failed: {e}\nserialized: {serialized}"));

    let latex_after = latex_of(&reparsed);
    assert_eq!(
        latex_before, latex_after,
        "LaTeX drift after OMML round-trip\nserialized: {serialized}"
    );
}

#[test]
fn latex_consistent_fraction_with_operators() {
    assert_latex_consistent(&[MathNode::Frac {
        numerator: vec![MathNode::Number(Cow::Borrowed("1"))],
        denominator: vec![
            text("x"),
            MathNode::Operator(Operator::Plus),
            MathNode::Number(Cow::Borrowed("2")),
        ],
        line_thickness: None,
        frac_type: None,
    }]);
}

#[test]
fn latex_consistent_root_of_power() {
    assert_latex_consistent(&[MathNode::Root {
        base: vec![MathNode::Power {
            base: vec![text("x")],
            exponent: vec![MathNode::Number(Cow::Borrowed("2"))],
        }],
        index: None,
    }]);
}
