use super::{flatten_text, number, roundtrip, roundtrip_text, single, text, write_bare};
use crate::ast::{Fence, FractionType, FunctionName, LargeOperator, MathNode, MatrixFence};
use crate::mtef::constants::*;
use std::borrow::Cow;

// ---------------------------------------------------------------------------
// Structures
// ---------------------------------------------------------------------------

#[test]
fn fractions_round_trip() {
    let node = MathNode::Frac {
        numerator: vec![number("1")],
        denominator: vec![text("x")],
        line_thickness: None,
        frac_type: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Frac {
            numerator,
            denominator,
            frac_type,
            ..
        } => {
            assert_eq!(flatten_text(numerator), "1");
            assert_eq!(flatten_text(denominator), "x");
            assert_eq!(*frac_type, None);
        },
        other => panic!("expected a fraction, got {other:?}"),
    });
}

#[test]
fn barless_fractions_keep_their_style() {
    let node = MathNode::Frac {
        numerator: vec![number("1")],
        denominator: vec![number("2")],
        line_thickness: None,
        frac_type: Some(FractionType::NoBar),
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Frac { frac_type, .. } => {
            assert_eq!(*frac_type, Some(FractionType::NoBar));
        },
        other => panic!("expected a fraction, got {other:?}"),
    });
}

#[test]
fn square_roots_round_trip() {
    let node = MathNode::Root {
        base: vec![text("x")],
        index: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Root { base, index } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(*index, None);
        },
        other => panic!("expected a root, got {other:?}"),
    });
}

#[test]
fn nth_roots_keep_radicand_and_degree_apart() {
    let node = MathNode::Root {
        base: vec![text("x")],
        index: Some(vec![number("3")]),
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Root { base, index } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(
                flatten_text(index.as_deref().expect("degree preserved")),
                "3"
            );
        },
        other => panic!("expected a root, got {other:?}"),
    });
}

#[test]
fn superscripts_round_trip() {
    let node = MathNode::Power {
        base: vec![text("x")],
        exponent: vec![number("2")],
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Power { base, exponent } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(flatten_text(exponent), "2");
        },
        other => panic!("expected a power, got {other:?}"),
    });
}

#[test]
fn subscripts_round_trip() {
    let node = MathNode::Sub {
        base: vec![text("a")],
        subscript: vec![text("i")],
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Sub { base, subscript } => {
            assert_eq!(flatten_text(base), "a");
            assert_eq!(flatten_text(subscript), "i");
        },
        other => panic!("expected a subscript, got {other:?}"),
    });
}

#[test]
fn sub_and_superscripts_round_trip_together() {
    let node = MathNode::SubSup {
        base: vec![text("x")],
        subscript: vec![text("i")],
        superscript: vec![number("2")],
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::SubSup {
            base,
            subscript,
            superscript,
        } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(flatten_text(subscript), "i");
            assert_eq!(flatten_text(superscript), "2");
        },
        other => panic!("expected a subsup, got {other:?}"),
    });
}

#[test]
fn scripts_attach_to_the_preceding_object_only() {
    // MathType scripts decorate one object, so material before the base stays
    // beside the script rather than inside it.
    let nodes = [
        text("f"),
        MathNode::Power {
            base: vec![text("x")],
            exponent: vec![number("2")],
        },
    ];

    roundtrip(&nodes, |recovered| {
        assert_eq!(recovered.len(), 2);
        assert_eq!(flatten_text(&recovered[..1]), "f");
        match &recovered[1] {
            MathNode::Power { base, .. } => assert_eq!(flatten_text(base), "x"),
            other => panic!("expected a power, got {other:?}"),
        }
    });
}

#[test]
fn underscripts_round_trip() {
    let node = MathNode::Under {
        base: vec![text("x")],
        under: vec![number("1")],
        position: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Under { base, under, .. } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(flatten_text(under), "1");
        },
        other => panic!("expected an under, got {other:?}"),
    });
}

#[test]
fn overscripts_round_trip() {
    let node = MathNode::Over {
        base: vec![text("x")],
        over: vec![number("2")],
        position: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Over { base, over, .. } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(flatten_text(over), "2");
        },
        other => panic!("expected an over, got {other:?}"),
    });
}

#[test]
fn under_and_overscripts_round_trip_together() {
    let node = MathNode::UnderOver {
        base: vec![text("x")],
        under: vec![number("1")],
        over: vec![number("2")],
        position: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::UnderOver {
            base, under, over, ..
        } => {
            assert_eq!(flatten_text(base), "x");
            assert_eq!(flatten_text(under), "1");
            assert_eq!(flatten_text(over), "2");
        },
        other => panic!("expected an underover, got {other:?}"),
    });
}

#[test]
fn every_fence_variant_round_trips() {
    // The delimiters MTEF has a template for come back unchanged; the rest map
    // onto the closest template MathType offers.
    let cases = [
        (Fence::Paren, Fence::Paren),
        (Fence::Bracket, Fence::Bracket),
        (Fence::Brace, Fence::Brace),
        (Fence::Angle, Fence::Angle),
        (Fence::Pipe, Fence::Pipe),
        (Fence::DoublePipe, Fence::DoublePipe),
        (Fence::Floor, Fence::Floor),
        (Fence::Ceiling, Fence::Ceiling),
        (Fence::SquareBracket, Fence::SquareBracket),
        (Fence::AngleBracket, Fence::Angle),
        (Fence::CurlyBrace, Fence::Brace),
    ];

    for (fence, expected) in cases {
        let node = MathNode::Fenced {
            open: fence,
            content: vec![text("x")],
            close: fence,
            separator: None,
        };
        roundtrip(&[node], |recovered| match single(recovered) {
            MathNode::Fenced {
                open,
                content,
                close,
                ..
            } => {
                assert_eq!(*open, expected, "opening delimiter for {fence:?}");
                assert_eq!(*close, expected, "closing delimiter for {fence:?}");
                assert_eq!(flatten_text(content), "x");
            },
            other => panic!("expected a fence for {fence:?}, got {other:?}"),
        });
    }
}

#[test]
fn unfenced_content_is_written_without_a_template() {
    let node = MathNode::Fenced {
        open: Fence::None,
        content: vec![text("x")],
        close: Fence::None,
        separator: None,
    };
    assert_eq!(roundtrip_text(&[node]), "x");
}

#[test]
fn one_sided_fences_pick_the_matching_variation() {
    let node = MathNode::Fenced {
        open: Fence::Brace,
        content: vec![text("x")],
        close: Fence::None,
        separator: None,
    };
    let bytes = write_bare(&[node]);

    // MTEF header (7), FULL, LINE + options, then the template record.
    assert_eq!(&bytes[10..14], &[TMPL, 0, TMPL_BRACE, TV_FENCE_LEFT as u8]);
}

#[test]
fn mismatched_delimiters_use_the_interval_template() {
    let node = MathNode::Fenced {
        open: Fence::Bracket,
        content: vec![text("x")],
        close: Fence::Paren,
        separator: None,
    };
    let bytes = write_bare(&[node]);

    // "[x)" is the LBRP form, variation 18 in the MTEF template table.
    assert_eq!(&bytes[10..14], &[TMPL, 0, TMPL_INTERVAL, 18]);
}

#[test]
fn large_operators_round_trip_with_limits() {
    let node = MathNode::LargeOp {
        operator: LargeOperator::Sum,
        lower_limit: Some(vec![text("i")]),
        upper_limit: Some(vec![text("n")]),
        integrand: Some(vec![text("x")]),
        hide_lower: false,
        hide_upper: false,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::LargeOp {
            operator,
            lower_limit,
            upper_limit,
            integrand,
            hide_lower,
            hide_upper,
        } => {
            assert_eq!(*operator, LargeOperator::Sum);
            assert_eq!(
                flatten_text(lower_limit.as_deref().expect("lower limit")),
                "i"
            );
            assert_eq!(
                flatten_text(upper_limit.as_deref().expect("upper limit")),
                "n"
            );
            assert_eq!(flatten_text(integrand.as_deref().expect("integrand")), "x");
            assert!(!hide_lower);
            assert!(!hide_upper);
        },
        other => panic!("expected a large operator, got {other:?}"),
    });
}

#[test]
fn the_large_operator_family_keeps_its_identity() {
    let cases = [
        LargeOperator::Integral,
        LargeOperator::DoubleIntegral,
        LargeOperator::TripleIntegral,
        LargeOperator::ContourIntegral,
        LargeOperator::SurfaceIntegral,
        LargeOperator::VolumeIntegral,
        LargeOperator::Sum,
        LargeOperator::Product,
        LargeOperator::Coproduct,
        LargeOperator::Union,
        LargeOperator::Intersection,
    ];

    for expected in cases {
        let node = MathNode::LargeOp {
            operator: expected,
            lower_limit: None,
            upper_limit: None,
            integrand: Some(vec![text("x")]),
            hide_lower: true,
            hide_upper: true,
        };
        roundtrip(&[node], |recovered| match single(recovered) {
            MathNode::LargeOp { operator, .. } => assert_eq!(*operator, expected),
            other => panic!("expected a large operator for {expected:?}, got {other:?}"),
        });
    }
}

#[test]
fn hidden_limits_are_not_written() {
    let node = MathNode::LargeOp {
        operator: LargeOperator::Integral,
        lower_limit: Some(vec![text("a")]),
        upper_limit: Some(vec![text("b")]),
        integrand: Some(vec![text("x")]),
        hide_lower: true,
        hide_upper: true,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::LargeOp {
            lower_limit,
            upper_limit,
            hide_lower,
            hide_upper,
            ..
        } => {
            assert_eq!(*lower_limit, None, "hidden limits carry no content");
            assert_eq!(*upper_limit, None);
            assert!(hide_lower);
            assert!(hide_upper);
        },
        other => panic!("expected a large operator, got {other:?}"),
    });
}

#[test]
fn word_operators_become_a_limit_over_the_operator_name() {
    let node = MathNode::LargeOp {
        operator: LargeOperator::Limit,
        lower_limit: Some(vec![text("n")]),
        upper_limit: None,
        integrand: Some(vec![text("x")]),
        hide_lower: false,
        hide_upper: true,
    };

    roundtrip(&[node], |recovered| {
        match &recovered[0] {
            MathNode::Under { base, under, .. } => {
                assert_eq!(flatten_text(base), "\\lim");
                assert_eq!(flatten_text(under), "n");
            },
            other => panic!("expected an under, got {other:?}"),
        }
        assert_eq!(
            flatten_text(&recovered[1..]),
            "x",
            "the body follows the operator"
        );
    });
}

#[test]
fn functions_round_trip_as_recognised_names() {
    let node = MathNode::Function {
        name: Cow::Borrowed("sin"),
        argument: vec![text("x")],
    };
    assert_eq!(roundtrip_text(&[node]), "\\sinx");
}

#[test]
fn predefined_functions_round_trip() {
    let node = MathNode::PredefinedFunction {
        function: FunctionName::Log,
        argument: vec![text("x")],
    };
    assert_eq!(roundtrip_text(&[node]), "\\logx");
}

#[test]
fn unknown_function_names_keep_their_letters() {
    let node = MathNode::Function {
        name: Cow::Borrowed("foo"),
        argument: vec![],
    };
    // The function typeface declares the run to be a function name, so an
    // unrecognised one stays a function rather than degrading to text. It must
    // not be spelled as the undefined command `\foo`.
    roundtrip(std::slice::from_ref(&node), |recovered| {
        assert_eq!(recovered, &[node.clone()][..]);
    });
    assert_eq!(roundtrip_text(&[node]), "\\operatorname{foo}");
}

#[test]
fn matrices_round_trip_cell_by_cell() {
    let node = MathNode::Matrix {
        rows: vec![
            vec![vec![number("1")], vec![number("2")]],
            vec![vec![number("3")], vec![number("4")]],
        ],
        fence_type: MatrixFence::None,
        properties: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Matrix { rows, .. } => {
            assert_eq!(rows.len(), 2);
            assert_eq!(rows[0].len(), 2);
            assert_eq!(flatten_text(&rows[0][0]), "1");
            assert_eq!(flatten_text(&rows[0][1]), "2");
            assert_eq!(flatten_text(&rows[1][0]), "3");
            assert_eq!(flatten_text(&rows[1][1]), "4");
        },
        other => panic!("expected a matrix, got {other:?}"),
    });
}

#[test]
fn bracketed_matrices_are_wrapped_in_a_fence() {
    let node = MathNode::Matrix {
        rows: vec![vec![vec![number("1")], vec![number("2")]]],
        fence_type: MatrixFence::Bracket,
        properties: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Fenced { open, content, .. } => {
            assert_eq!(*open, Fence::Bracket);
            assert!(matches!(content.first(), Some(MathNode::Matrix { .. })));
        },
        other => panic!("expected a fenced matrix, got {other:?}"),
    });
}

#[test]
fn ragged_matrix_rows_are_padded() {
    let node = MathNode::Matrix {
        rows: vec![
            vec![vec![number("1")], vec![number("2")]],
            vec![vec![number("3")]],
        ],
        fence_type: MatrixFence::None,
        properties: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Matrix { rows, .. } => {
            assert_eq!(rows[1].len(), 2);
            assert_eq!(flatten_text(&rows[1][1]), "");
        },
        other => panic!("expected a matrix, got {other:?}"),
    });
}

#[test]
fn equation_arrays_become_a_pile_of_rows() {
    let node = MathNode::EqArray {
        rows: vec![vec![text("x")], vec![text("y")]],
        properties: None,
    };

    // A pile of lines is the closest MTEF construct; the reader reports it as a
    // single-column matrix.
    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Matrix { rows, .. } => {
            assert_eq!(rows.len(), 2);
            assert_eq!(flatten_text(&rows[0][0]), "x");
            assert_eq!(flatten_text(&rows[1][0]), "y");
        },
        other => panic!("expected a matrix, got {other:?}"),
    });
}

#[test]
fn line_breaks_split_the_equation_into_a_pile() {
    let nodes = [text("x"), MathNode::LineBreak, text("y")];

    roundtrip(&nodes, |recovered| match single(recovered) {
        MathNode::Matrix { rows, .. } => {
            assert_eq!(rows.len(), 2);
            assert_eq!(flatten_text(&rows[0][0]), "x");
            assert_eq!(flatten_text(&rows[1][0]), "y");
        },
        other => panic!("expected a pile, got {other:?}"),
    });
}
