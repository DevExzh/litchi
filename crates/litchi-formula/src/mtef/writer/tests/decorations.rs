use super::{flatten_text, number, roundtrip, roundtrip_text, single, text, write_bare};
use crate::ast::MatrixFence;
use crate::ast::{AccentType, LineStyle, MathNode, Operator, Position, StyleType};
use crate::mtef::Limits;
use crate::mtef::writer::{MtefWriteError, MtefWriter};
use std::borrow::Cow;

// ---------------------------------------------------------------------------
// Decorations, runs and degradation
// ---------------------------------------------------------------------------

#[test]
fn accents_become_character_embellishments() {
    let cases = [
        (AccentType::Hat, "\\hat{x} "),
        (AccentType::Tilde, "\\tilde{x} "),
        (AccentType::Dot, "\\dot{x} "),
        (AccentType::DoubleDot, "\\ddot{x} "),
        (AccentType::Bar, "\\bar{x} "),
        (AccentType::Vec, "\\vec{x} "),
    ];

    for (accent, expected) in cases {
        let node = MathNode::Accent {
            base: Box::new(vec![text("x")]),
            accent,
            position: None,
        };
        assert_eq!(roundtrip_text(&[node]), expected, "{accent:?}");
    }
}

#[test]
fn an_embellished_character_does_not_truncate_its_line() {
    // The embellishment list is terminated by its own END record; mistaking it
    // for the line's terminator would drop everything after the accent.
    let nodes = [
        MathNode::Accent {
            base: Box::new(vec![text("x")]),
            accent: AccentType::Hat,
            position: None,
        },
        MathNode::Operator(Operator::Plus),
        text("y"),
    ];
    assert_eq!(roundtrip_text(&nodes), "\\hat{x} +y");
}

#[test]
fn accents_over_structures_degrade_to_the_structure() {
    // An embellishment decorates a character, so there is nothing to attach to.
    let node = MathNode::Accent {
        base: Box::new(vec![MathNode::Frac {
            numerator: vec![number("1")],
            denominator: vec![number("2")],
            line_thickness: None,
            frac_type: None,
        }]),
        accent: AccentType::Hat,
        position: None,
    };

    roundtrip(&[node], |recovered| {
        assert!(matches!(single(recovered), MathNode::Frac { .. }));
    });
}

#[test]
fn bars_become_an_overbar_template() {
    let node = MathNode::Bar {
        base: Box::new(vec![text("xy")]),
        position: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Run {
            content, overline, ..
        } => {
            assert_eq!(*overline, Some(LineStyle::Single));
            assert_eq!(flatten_text(content), "xy");
        },
        other => panic!("expected an overlined run, got {other:?}"),
    });
}

#[test]
fn group_characters_round_trip_with_their_position() {
    for position in [Position::Top, Position::Bottom] {
        let node = MathNode::GroupChar {
            base: Box::new(vec![text("x")]),
            character: None,
            position: Some(position),
            vertical_alignment: None,
        };
        roundtrip(&[node], |recovered| match single(recovered) {
            MathNode::GroupChar {
                base,
                character,
                position: recovered_position,
                ..
            } => {
                assert_eq!(flatten_text(base), "x");
                assert_eq!(*recovered_position, Some(position));
                assert!(character.is_some(), "the brace character is restored");
            },
            other => panic!("expected a group character, got {other:?}"),
        });
    }
}

#[test]
fn runs_round_trip_their_rules() {
    let node = MathNode::Run {
        content: vec![text("x")],
        literal: None,
        style: None,
        font: None,
        color: None,
        underline: Some(LineStyle::Single),
        overline: None,
        strike_through: None,
        double_strike_through: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Run {
            content, underline, ..
        } => {
            assert_eq!(*underline, Some(LineStyle::Single));
            assert_eq!(flatten_text(content), "x");
        },
        other => panic!("expected an underlined run, got {other:?}"),
    });
}

#[test]
fn ruled_runs_match_single_and_double_encoded_depth() {
    let single = MathNode::Run {
        content: vec![text("x")],
        literal: None,
        style: None,
        font: None,
        color: None,
        underline: Some(LineStyle::Single),
        overline: None,
        strike_through: None,
        double_strike_through: None,
    };
    let double = MathNode::Run {
        content: vec![text("x")],
        literal: None,
        style: None,
        font: None,
        color: None,
        underline: Some(LineStyle::Single),
        overline: Some(LineStyle::Single),
        strike_through: None,
        double_strike_through: None,
    };

    let depth_three = Limits::new(4096, 3, 256, 256, 4096, 4096).unwrap();
    MtefWriter::new(depth_three)
        .write_nodes(&[single])
        .expect("root line plus one ruled template and slot line fit");
    let depth_four = Limits::new(4096, 4, 256, 256, 4096, 4096).unwrap();
    assert!(matches!(
        MtefWriter::new(depth_four).write_nodes(std::slice::from_ref(&double)),
        Err(MtefWriteError::LimitExceeded {
            kind: crate::mtef::LimitKind::Depth,
            actual: 5,
            ..
        })
    ));

    let depth_five = Limits::new(4096, 5, 256, 256, 4096, 4096).unwrap();
    MtefWriter::new(depth_five)
        .write_nodes(&[double])
        .expect("root line plus two ruled templates and slot lines fit at depth five");
}

#[test]
fn runs_can_name_a_font() {
    let font_run = || MathNode::Run {
        content: vec![text("x")],
        literal: None,
        style: None,
        font: Some(Cow::Borrowed("Times New Roman")),
        color: None,
        underline: None,
        overline: None,
        strike_through: None,
        double_strike_through: None,
    };

    let bytes = write_bare(&[font_run()]);
    assert!(
        bytes.windows(15).any(|window| window == b"Times New Roman"),
        "the font name is written"
    );
    assert_eq!(
        roundtrip_text(&[font_run()]),
        "x",
        "a font record does not disturb the content"
    );
}

#[test]
fn bold_styles_select_the_bold_typeface() {
    let node = MathNode::Style {
        style: StyleType::Bold,
        content: vec![text("x")],
    };
    assert_eq!(roundtrip_text(&[node]), "\\mathbf{x}");
}

#[test]
fn phantoms_and_border_boxes_degrade_to_their_content() {
    let phantom = MathNode::Phantom(Box::new(vec![text("x")]));
    assert_eq!(roundtrip_text(&[phantom]), "x");

    let border_box = MathNode::BorderBox {
        content: Box::new(vec![text("y")]),
        style: None,
    };
    assert_eq!(roundtrip_text(&[border_box]), "y");
}

#[test]
fn wrapper_nodes_are_transparent() {
    let node = MathNode::Frac {
        numerator: vec![MathNode::Numerator(Box::new(vec![number("1")]))],
        denominator: vec![MathNode::Denominator(Box::new(vec![number("2")]))],
        line_thickness: None,
        frac_type: None,
    };

    roundtrip(&[node], |recovered| match single(recovered) {
        MathNode::Frac {
            numerator,
            denominator,
            ..
        } => {
            assert_eq!(flatten_text(numerator), "1");
            assert_eq!(flatten_text(denominator), "2");
        },
        other => panic!("expected a fraction, got {other:?}"),
    });
}

#[test]
fn rows_are_flattened_into_the_enclosing_line() {
    let node = MathNode::Row(vec![
        text("x"),
        MathNode::Operator(Operator::Plus),
        text("y"),
    ]);
    assert_eq!(roundtrip_text(&[node]), "x+y");
}

// ---------------------------------------------------------------------------
// Limits
// ---------------------------------------------------------------------------

#[test]
fn deeply_nested_formulas_are_rejected_rather_than_overflowing() {
    let mut node = text("x");
    for _ in 0..200 {
        node = MathNode::Frac {
            numerator: vec![node],
            denominator: vec![number("2")],
            line_thickness: None,
            frac_type: None,
        };
    }

    assert!(matches!(
        MtefWriter::new(Limits::DEFAULT).write_nodes(&[node]),
        Err(MtefWriteError::LimitExceeded {
            kind: crate::mtef::LimitKind::Depth,
            ..
        })
    ));
}

#[test]
fn preflight_matches_line_break_framing_at_the_output_boundary() {
    // A bare stream containing one break is exactly 20 bytes: the MTEF
    // header/full framing (11), a PILE with two empty LINEs (8), and the
    // outer END (1).
    let exact = Limits::new(4096, 8, 256, 256, 4096, 20).unwrap();
    let bytes = MtefWriter::with_options(
        exact,
        crate::mtef::writer::MtefWriteOptions::without_ole_header(),
    )
    .write_nodes(&[MathNode::LineBreak])
    .expect("the exact output boundary is admitted");
    assert_eq!(bytes.len(), 20);

    let under = Limits::new(4096, 8, 256, 256, 4096, 19).unwrap();
    assert!(matches!(
        MtefWriter::with_options(
            under,
            crate::mtef::writer::MtefWriteOptions::without_ole_header(),
        )
        .write_nodes(&[MathNode::LineBreak]),
        Err(MtefWriteError::LimitExceeded {
            kind: crate::mtef::LimitKind::OutputBytes,
            actual: 20,
            ..
        })
    ));
}

#[test]
fn preflight_and_emitter_share_the_exact_depth_boundary() {
    let exact = Limits::new(4096, 1, 256, 256, 4096, 4096).unwrap();
    MtefWriter::new(exact)
        .write_nodes(&[text("x")])
        .expect("a leaf fits at depth one");
    MtefWriter::new(exact)
        .write_nodes(&[MathNode::Row(vec![text("x")])])
        .expect("transparent rows do not consume encoded depth");
    assert!(matches!(
        MtefWriter::new(exact).write_nodes(&[MathNode::LineBreak]),
        Err(MtefWriteError::LimitExceeded {
            kind: crate::mtef::LimitKind::Depth,
            actual: 2,
            ..
        })
    ));
    let pile_exact = Limits::new(4096, 2, 256, 256, 4096, 4096).unwrap();
    MtefWriter::new(pile_exact)
        .write_nodes(&[MathNode::LineBreak])
        .expect("a root pile and its row fit at depth two");

    let nested = MathNode::Frac {
        numerator: vec![text("x")],
        denominator: vec![text("y")],
        line_thickness: None,
        frac_type: None,
    };
    let nested_under = Limits::new(4096, 2, 256, 256, 4096, 4096).unwrap();
    assert!(matches!(
        MtefWriter::new(nested_under).write_nodes(&[nested]),
        Err(MtefWriteError::LimitExceeded {
            kind: crate::mtef::LimitKind::Depth,
            actual: 3,
            ..
        })
    ));
}

#[test]
fn nested_templates_and_matrix_cells_use_their_encoded_line_depth() {
    let fraction = MathNode::Frac {
        numerator: vec![text("x")],
        denominator: vec![text("y")],
        line_thickness: None,
        frac_type: None,
    };
    let root = MathNode::Root {
        base: vec![text("x")],
        index: None,
    };
    let matrix = MathNode::Matrix {
        rows: vec![vec![vec![text("x")]]],
        fence_type: MatrixFence::None,
        properties: None,
    };

    let under = Limits::new(4096, 2, 256, 256, 4096, 4096).unwrap();
    for node in [fraction.clone(), root.clone(), matrix.clone()] {
        assert!(matches!(
            MtefWriter::new(under).write_nodes(&[node]),
            Err(MtefWriteError::LimitExceeded {
                kind: crate::mtef::LimitKind::Depth,
                actual: 3,
                ..
            })
        ));
    }

    let exact = Limits::new(4096, 3, 256, 256, 4096, 4096).unwrap();
    for node in [fraction, root, matrix] {
        MtefWriter::new(exact)
            .write_nodes(&[node])
            .expect("three encoded levels are sufficient");
    }
}

#[test]
fn oversized_matrices_are_rejected() {
    let row: Vec<Vec<MathNode<'static>>> = (0..64).map(|_| vec![number("1")]).collect();
    let node = MathNode::Matrix {
        rows: vec![row],
        fence_type: MatrixFence::None,
        properties: None,
    };

    assert!(matches!(
        MtefWriter::new(Limits::DEFAULT).write_nodes(&[node]),
        Err(MtefWriteError::MatrixTooLarge { .. })
    ));
}
