// Tokenizing, grouping, scripts, fractions and fences

use super::*;

// -- literals ------------------------------------------------------------

#[test]
fn parses_each_letter_as_its_own_identifier() {
    assert_eq!(parse("abc"), vec![text("a"), text("b"), text("c")]);
}

#[test]
fn merges_digit_runs_into_a_single_number() {
    assert_eq!(parse("1234"), vec![number("1234")]);
    assert_eq!(parse("3.14"), vec![number("3.14")]);
    assert_eq!(parse("1 2"), vec![number("1"), number("2")]);
}

#[test]
fn borrows_literals_from_the_input_without_allocating() {
    let input = String::from("xy12");
    let nodes = LatexParser::new().parse(&input).expect("input parses");
    for node in &nodes {
        let borrowed = match node {
            MathNode::Text(value) | MathNode::Number(value) => value,
            other => panic!("unexpected node {other:?}"),
        };
        assert!(matches!(borrowed, Cow::Borrowed(_)));
    }
}

#[test]
fn a_trailing_decimal_point_is_not_part_of_the_number() {
    assert_eq!(parse("1."), vec![number("1"), text(".")]);
}

#[test]
fn maps_ascii_operators_that_render_identically() {
    assert_eq!(
        parse("a+b-c=d<e>f"),
        vec![
            text("a"),
            MathNode::Operator(Operator::Plus),
            text("b"),
            MathNode::Operator(Operator::Minus),
            text("c"),
            MathNode::Operator(Operator::Equals),
            text("d"),
            MathNode::Operator(Operator::LessThan),
            text("e"),
            MathNode::Operator(Operator::GreaterThan),
            text("f"),
        ]
    );
}

#[test]
fn keeps_punctuation_without_an_exact_operator_as_text() {
    assert_eq!(
        parse("!,*/"),
        vec![text("!"), text(","), text("*"), text("/")]
    );
}

#[test]
fn parses_a_prime_as_an_operator() {
    assert_eq!(
        parse("x'"),
        vec![text("x"), MathNode::Operator(Operator::Prime)]
    );
}

#[test]
fn passes_unicode_through_as_a_symbol() {
    let MathNode::Symbol(symbol) = parse_one("α") else {
        panic!("expected a symbol");
    };
    assert_eq!(symbol.name.as_ref(), "α");
    assert_eq!(symbol.unicode, Some('α'));
}

#[test]
fn strips_math_mode_delimiters() {
    assert_eq!(parse("$x$"), vec![text("x")]);
    assert_eq!(parse("\\(x\\)"), vec![text("x")]);
    assert_eq!(parse("\\[x\\]"), vec![text("x")]);
}

#[test]
fn strips_comments() {
    assert_eq!(parse("a% a comment\nb"), vec![text("a"), text("b")]);
}

// -- groups and scripts --------------------------------------------------

#[test]
fn inlines_a_group_that_adds_no_structure() {
    assert_eq!(parse("{a}"), vec![text("a")]);
    assert_eq!(parse("{}"), Vec::new());
}

#[test]
fn keeps_a_multi_element_group_as_a_row() {
    assert_eq!(
        parse("{a+b}"),
        vec![MathNode::Row(vec![
            text("a"),
            MathNode::Operator(Operator::Plus),
            text("b"),
        ])]
    );
}

#[test]
fn binds_a_superscript_to_the_preceding_atom() {
    assert_eq!(
        parse("ab^2"),
        vec![
            text("a"),
            MathNode::Power {
                base: vec![text("b")],
                exponent: vec![number("2")],
            }
        ]
    );
}

#[test]
fn binds_a_script_to_a_whole_preceding_group() {
    assert_eq!(
        parse("{a+b}^2"),
        vec![MathNode::Power {
            base: vec![text("a"), MathNode::Operator(Operator::Plus), text("b")],
            exponent: vec![number("2")],
        }]
    );
}

#[test]
fn combines_a_subscript_and_superscript_in_either_order() {
    let expected = MathNode::SubSup {
        base: vec![text("x")],
        subscript: vec![text("i")],
        superscript: vec![number("2")],
    };
    assert_eq!(parse_one("x_i^2"), expected);
    assert_eq!(parse_one("x^2_i"), expected);
}

#[test]
fn takes_only_one_digit_for_an_unbraced_script() {
    assert_eq!(
        parse("x^12"),
        vec![
            MathNode::Power {
                base: vec![text("x")],
                exponent: vec![number("1")],
            },
            number("2"),
        ]
    );
}

#[test]
fn parses_prescripts_written_with_an_empty_group() {
    assert_eq!(
        parse_one("{}_a^b X"),
        MathNode::PreSubSup {
            base: vec![text("X")],
            pre_subscript: vec![text("a")],
            pre_superscript: vec![text("b")],
        }
    );
    assert_eq!(
        parse_one("{}_a X"),
        MathNode::PreSub {
            base: vec![text("X")],
            pre_subscript: vec![text("a")],
        }
    );
    assert_eq!(
        parse_one("{}^b X"),
        MathNode::PreSup {
            base: vec![text("X")],
            pre_superscript: vec![text("b")],
        }
    );
}

// -- fractions and roots -------------------------------------------------

#[test]
fn parses_every_fraction_spelling() {
    let expected = MathNode::Frac {
        numerator: vec![text("a")],
        denominator: vec![text("b")],
        line_thickness: None,
        frac_type: Some(FractionType::Bar),
    };
    for input in [
        "\\frac{a}{b}",
        "\\dfrac{a}{b}",
        "\\tfrac{a}{b}",
        "\\frac ab",
    ] {
        assert_eq!(parse_one(input), expected, "`{input}` should be a fraction");
    }
}

#[test]
fn parses_binomials_as_a_parenthesised_barless_fraction() {
    let expected = MathNode::Fenced {
        open: Fence::Paren,
        content: vec![MathNode::Frac {
            numerator: vec![text("n")],
            denominator: vec![text("k")],
            line_thickness: None,
            frac_type: Some(FractionType::NoBar),
        }],
        close: Fence::Paren,
        separator: None,
    };
    assert_eq!(parse_one("\\binom{n}{k}"), expected);
    assert_eq!(parse_one("{n \\choose k}"), expected);
}

#[test]
fn parses_the_infix_over_command() {
    assert_eq!(
        parse_one("{a \\over b}"),
        MathNode::Frac {
            numerator: vec![text("a")],
            denominator: vec![text("b")],
            line_thickness: None,
            frac_type: Some(FractionType::Bar),
        }
    );
}

#[test]
fn parses_square_roots_with_and_without_an_index() {
    assert_eq!(
        parse_one("\\sqrt{x}"),
        MathNode::Root {
            base: vec![text("x")],
            index: None,
        }
    );
    assert_eq!(
        parse_one("\\sqrt[3]{x}"),
        MathNode::Root {
            base: vec![text("x")],
            index: Some(vec![number("3")]),
        }
    );
}

// -- fences --------------------------------------------------------------

#[test]
fn parses_explicit_left_right_pairs() {
    for (input, open, close) in [
        ("\\left( x \\right)", Fence::Paren, Fence::Paren),
        ("\\left[ x \\right]", Fence::Bracket, Fence::Bracket),
        ("\\left\\{ x \\right\\}", Fence::Brace, Fence::Brace),
        ("\\left| x \\right|", Fence::Pipe, Fence::Pipe),
        (
            "\\left\\| x \\right\\|",
            Fence::DoublePipe,
            Fence::DoublePipe,
        ),
        (
            "\\left\\langle x \\right\\rangle",
            Fence::Angle,
            Fence::Angle,
        ),
        (
            "\\left\\lfloor x \\right\\rfloor",
            Fence::Floor,
            Fence::Floor,
        ),
        (
            "\\left\\lceil x \\right\\rceil",
            Fence::Ceiling,
            Fence::Ceiling,
        ),
        ("\\left. x \\right)", Fence::None, Fence::Paren),
        ("\\left( x \\right.", Fence::Paren, Fence::None),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::Fenced {
                open,
                content: vec![text("x")],
                close,
                separator: None,
            },
            "`{input}` should fence correctly"
        );
    }
}

#[test]
fn parses_bare_delimiter_pairs() {
    assert_eq!(
        parse_one("(x)"),
        MathNode::Fenced {
            open: Fence::Paren,
            content: vec![text("x")],
            close: Fence::Paren,
            separator: None,
        }
    );
    assert_eq!(
        parse_one("\\langle x \\rangle"),
        MathNode::Fenced {
            open: Fence::Angle,
            content: vec![text("x")],
            close: Fence::Angle,
            separator: None,
        }
    );
}

#[test]
fn nests_bare_delimiter_pairs() {
    let MathNode::Fenced { content, .. } = parse_one("((x))") else {
        panic!("expected a fenced node");
    };
    assert!(matches!(content.as_slice(), [MathNode::Fenced { .. }]));
}

#[test]
fn an_unpaired_delimiter_stays_an_ordinary_atom() {
    assert_eq!(parse("(x"), vec![text("("), text("x")]);
}
