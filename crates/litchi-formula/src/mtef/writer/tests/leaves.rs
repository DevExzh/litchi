use super::{number, roundtrip_text, text};
use crate::ast::{MathNode, Operator, PredefinedSymbol, SpaceType, Symbol};
use crate::mtef::writer::{MtefWriteError, MtefWriter};
use std::borrow::Cow;

// ---------------------------------------------------------------------------
// Leaves
// ---------------------------------------------------------------------------

#[test]
fn text_and_numbers_round_trip() {
    assert_eq!(roundtrip_text(&[text("xy"), number("12")]), "xy12");
}

#[test]
fn operators_round_trip_through_the_symbol_typeface() {
    let nodes = [
        MathNode::Operator(Operator::Equals),
        MathNode::Operator(Operator::Plus),
        MathNode::Operator(Operator::LessThanOrEqual),
    ];
    assert_eq!(roundtrip_text(&nodes), "=+\\leq ");
}

#[test]
fn predefined_symbols_round_trip() {
    let nodes = [
        MathNode::PredefinedSymbol(PredefinedSymbol::Alpha),
        MathNode::PredefinedSymbol(PredefinedSymbol::OmegaCap),
    ];
    assert_eq!(roundtrip_text(&nodes), "\\alpha \\Omega ");
}

#[test]
fn unicode_symbols_round_trip() {
    let symbol = MathNode::Symbol(Symbol {
        name: Cow::Borrowed("infinity"),
        unicode: Some('∞'),
        variant: None,
    });
    assert_eq!(roundtrip_text(&[symbol]), "\\infty ");
}

#[test]
fn named_symbols_without_a_codepoint_fall_back_to_their_name() {
    let symbol = MathNode::Symbol(Symbol {
        name: Cow::Borrowed("ab"),
        unicode: None,
        variant: None,
    });
    assert_eq!(roundtrip_text(&[symbol]), "ab");
}

#[test]
fn spaces_round_trip_as_spacing_commands() {
    let nodes = [
        MathNode::Space(SpaceType::Thin),
        MathNode::Space(SpaceType::Quad),
    ];
    assert_eq!(roundtrip_text(&nodes), "\\,\\quad ");
}

#[test]
fn error_nodes_are_dropped() {
    let nodes = [MathNode::Error(Cow::Borrowed("bad")), text("x")];
    assert_eq!(roundtrip_text(&nodes), "x");
}

#[test]
fn characters_outside_the_bmp_are_rejected() {
    assert_eq!(
        MtefWriter::new().write_nodes(&[text("𝕏")]),
        Err(MtefWriteError::UnsupportedCharacter('𝕏'))
    );
}
