// Command vocabulary: operators, functions, accents, styles and environments

use super::*;

// -- large operators -----------------------------------------------------

#[test]
fn parses_large_operators_with_both_limits() {
    assert_eq!(
        parse_one("\\sum_{i}^{n}"),
        MathNode::LargeOp {
            operator: LargeOperator::Sum,
            lower_limit: Some(vec![text("i")]),
            upper_limit: Some(vec![text("n")]),
            integrand: None,
            hide_lower: false,
            hide_upper: false,
        }
    );
}

#[test]
fn hides_limits_that_were_not_written() {
    let MathNode::LargeOp {
        hide_lower,
        hide_upper,
        ..
    } = parse_one("\\int_0")
    else {
        panic!("expected a large operator");
    };
    assert!(!hide_lower);
    assert!(hide_upper);
}

#[test]
fn accepts_limits_annotations() {
    for input in ["\\prod\\limits_{i}", "\\prod\\nolimits_{i}"] {
        let MathNode::LargeOp {
            operator,
            lower_limit,
            ..
        } = parse_one(input)
        else {
            panic!("`{input}` should be a large operator");
        };
        assert_eq!(operator, LargeOperator::Product);
        assert_eq!(lower_limit, Some(vec![text("i")]));
    }
}

#[test]
fn covers_the_whole_large_operator_vocabulary() {
    for (input, expected) in [
        ("\\sum", LargeOperator::Sum),
        ("\\prod", LargeOperator::Product),
        ("\\coprod", LargeOperator::Coproduct),
        ("\\int", LargeOperator::Integral),
        ("\\iint", LargeOperator::DoubleIntegral),
        ("\\iiint", LargeOperator::TripleIntegral),
        ("\\oint", LargeOperator::ContourIntegral),
        ("\\bigcup", LargeOperator::Union),
        ("\\bigcap", LargeOperator::Intersection),
        ("\\lim", LargeOperator::Limit),
    ] {
        let MathNode::LargeOp { operator, .. } = parse_one(input) else {
            panic!("`{input}` should be a large operator");
        };
        assert_eq!(operator, expected);
    }
}

// -- functions -----------------------------------------------------------

#[test]
fn parses_named_functions_with_a_predefined_variant() {
    for (input, expected) in [
        ("\\sin x", FunctionName::Sin),
        ("\\cos x", FunctionName::Cos),
        ("\\tan x", FunctionName::Tan),
        ("\\sec x", FunctionName::Sec),
        ("\\csc x", FunctionName::Csc),
        ("\\cot x", FunctionName::Cot),
        ("\\arcsin x", FunctionName::ArcSin),
        ("\\arccos x", FunctionName::ArcCos),
        ("\\arctan x", FunctionName::ArcTan),
        ("\\sinh x", FunctionName::Sinh),
        ("\\cosh x", FunctionName::Cosh),
        ("\\tanh x", FunctionName::Tanh),
        ("\\log x", FunctionName::Log),
        ("\\ln x", FunctionName::Ln),
        ("\\exp x", FunctionName::Exp),
        ("\\min x", FunctionName::Min),
        ("\\max x", FunctionName::Max),
        ("\\sup x", FunctionName::Sup),
        ("\\inf x", FunctionName::Inf),
        ("\\det x", FunctionName::Det),
        ("\\dim x", FunctionName::Dim),
        ("\\ker x", FunctionName::Ker),
        ("\\gcd x", FunctionName::Gcd),
        ("\\arg x", FunctionName::Arg),
        ("\\bmod x", FunctionName::Mod),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::PredefinedFunction {
                function: expected,
                argument: vec![text("x")],
            },
            "`{input}` should be a predefined function"
        );
    }
}

#[test]
fn parses_named_functions_without_a_predefined_variant() {
    for name in ["deg", "Pr"] {
        let input = format!("\\{name} x");
        assert_eq!(
            LatexParser::new().parse(&input).expect("parses"),
            vec![MathNode::Function {
                name: Cow::Borrowed(name),
                argument: vec![text("x")],
            }]
        );
    }
}

#[test]
fn a_function_does_not_swallow_a_following_operator() {
    assert_eq!(
        parse("\\sin + 1"),
        vec![
            MathNode::PredefinedFunction {
                function: FunctionName::Sin,
                argument: Vec::new(),
            },
            MathNode::Operator(Operator::Plus),
            number("1"),
        ]
    );
}

#[test]
fn parses_operatorname() {
    assert_eq!(
        parse_one("\\operatorname{argmax}{x}"),
        MathNode::Function {
            name: Cow::Borrowed("argmax"),
            argument: vec![text("x")],
        }
    );
}

// -- accents, bars and braces --------------------------------------------

#[test]
fn parses_every_accent() {
    for (input, expected) in [
        ("\\hat{x}", AccentType::Hat),
        ("\\widehat{x}", AccentType::Hat),
        ("\\tilde{x}", AccentType::Tilde),
        ("\\widetilde{x}", AccentType::Tilde),
        ("\\bar{x}", AccentType::Bar),
        ("\\vec{x}", AccentType::Vec),
        ("\\dot{x}", AccentType::Dot),
        ("\\ddot{x}", AccentType::DoubleDot),
        ("\\dddot{x}", AccentType::TripleDot),
        ("\\acute{x}", AccentType::Acute),
        ("\\grave{x}", AccentType::Grave),
        ("\\check{x}", AccentType::Check),
        ("\\breve{x}", AccentType::Breve),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::Accent {
                base: Box::new(vec![text("x")]),
                accent: expected,
                position: Some(Position::Top),
            },
            "`{input}` should be an accent"
        );
    }
}

#[test]
fn parses_overline_and_underline() {
    assert_eq!(
        parse_one("\\overline{x}"),
        MathNode::Bar {
            base: Box::new(vec![text("x")]),
            position: Some(Position::Top),
        }
    );
    assert_eq!(
        parse_one("\\underline{x}"),
        MathNode::Under {
            base: vec![text("x")],
            under: Vec::new(),
            position: Some(Position::Bottom),
        }
    );
}

#[test]
fn parses_over_and_under_braces() {
    let MathNode::GroupChar { position, .. } = parse_one("\\overbrace{x}") else {
        panic!("expected a group character");
    };
    assert_eq!(position, Some(Position::Top));

    let MathNode::GroupChar { position, .. } = parse_one("\\underbrace{x}") else {
        panic!("expected a group character");
    };
    assert_eq!(position, Some(Position::Bottom));
}

#[test]
fn parses_under_and_over_sets() {
    assert_eq!(
        parse_one("\\underset{u}{b}"),
        MathNode::Under {
            base: vec![text("b")],
            under: vec![text("u")],
            position: None,
        }
    );
    assert_eq!(
        parse_one("\\overset{o}{b}"),
        MathNode::Over {
            base: vec![text("b")],
            over: vec![text("o")],
            position: None,
        }
    );
}

// -- styles, boxes and spacing -------------------------------------------

#[test]
fn parses_math_alphabets() {
    for (input, expected) in [
        ("\\mathrm{x}", StyleType::Normal),
        ("\\mathbf{x}", StyleType::Bold),
        ("\\mathit{x}", StyleType::Italic),
        ("\\mathbb{x}", StyleType::DoubleStruck),
        ("\\mathcal{x}", StyleType::Script),
        ("\\mathfrak{x}", StyleType::Fraktur),
        ("\\mathsf{x}", StyleType::SansSerif),
        ("\\mathtt{x}", StyleType::Monospace),
        ("\\boldsymbol{x}", StyleType::BoldItalic),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::Style {
                style: expected,
                content: vec![text("x")],
            },
            "`{input}` should select an alphabet"
        );
    }
}

#[test]
fn keeps_text_arguments_verbatim() {
    let MathNode::Run { content, style, .. } = parse_one("\\text{a b}") else {
        panic!("expected a literal run");
    };
    assert_eq!(content, vec![text("a b")]);
    assert_eq!(style, None);

    let MathNode::Run { style, .. } = parse_one("\\textbf{a b}") else {
        panic!("expected a literal run");
    };
    assert_eq!(style, Some(StyleType::Bold));
}

#[test]
fn unescapes_text_arguments() {
    let MathNode::Run { content, .. } = parse_one("\\text{a\\ b\\%c}") else {
        panic!("expected a literal run");
    };
    assert_eq!(content, vec![text("a b%c")]);
}

#[test]
fn parses_phantoms_and_boxes() {
    assert_eq!(
        parse_one("\\phantom{x}"),
        MathNode::Phantom(Box::new(vec![text("x")]))
    );
    for input in ["\\boxed{x}", "\\fbox{x}"] {
        assert_eq!(
            LatexParser::new().parse(input).expect("parses"),
            vec![MathNode::BorderBox {
                content: Box::new(vec![text("x")]),
                style: None,
            }]
        );
    }
}

#[test]
fn parses_every_space_command() {
    for (input, expected) in [
        ("\\,", SpaceType::Thin),
        ("\\:", SpaceType::Medium),
        ("\\;", SpaceType::Thick),
        ("\\!", SpaceType::Negative),
        ("\\quad", SpaceType::Quad),
        ("\\qquad", SpaceType::QQuad),
        ("\\ ", SpaceType::Medium),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::Space(expected),
            "`{input}` should be a space"
        );
    }
}

#[test]
fn parses_a_line_break() {
    assert_eq!(parse_one("\\\\"), MathNode::LineBreak);
    assert_eq!(parse_one("\\\\[6pt]"), MathNode::LineBreak);
}

// -- environments --------------------------------------------------------

#[test]
fn parses_matrix_environments_with_the_matching_fence() {
    for (name, fence) in [
        ("matrix", MatrixFence::None),
        ("pmatrix", MatrixFence::Paren),
        ("bmatrix", MatrixFence::Bracket),
        ("Bmatrix", MatrixFence::Brace),
        ("vmatrix", MatrixFence::Pipe),
        ("Vmatrix", MatrixFence::DoublePipe),
        ("smallmatrix", MatrixFence::None),
        ("cases", MatrixFence::Brace),
    ] {
        let input = format!("\\begin{{{name}}}a & b \\\\ c & d\\end{{{name}}}");
        let nodes = LatexParser::new()
            .parse(&input)
            .expect("environment parses");
        assert_eq!(
            nodes,
            vec![MathNode::Matrix {
                rows: vec![
                    vec![vec![text("a")], vec![text("b")]],
                    vec![vec![text("c")], vec![text("d")]],
                ],
                fence_type: fence,
                properties: None,
            }],
            "`{name}` should map to {fence:?}"
        );
    }
}

#[test]
fn consumes_and_ignores_an_array_column_specification() {
    assert_eq!(
        parse_one("\\begin{array}{l|cr}a & b\\end{array}"),
        MathNode::Matrix {
            rows: vec![vec![vec![text("a")], vec![text("b")]]],
            fence_type: MatrixFence::None,
            properties: None,
        }
    );
}

#[test]
fn parses_alignment_environments_as_equation_arrays() {
    for name in ["aligned", "align", "align*", "gathered", "gather", "split"] {
        let input = format!("\\begin{{{name}}}a \\\\ b\\end{{{name}}}");
        assert_eq!(
            LatexParser::new()
                .parse(&input)
                .expect("environment parses"),
            vec![MathNode::EqArray {
                rows: vec![vec![text("a")], vec![text("b")]],
                properties: None,
            }],
            "`{name}` should map to an equation array"
        );
    }
}

#[test]
fn joins_alignment_cells_dropping_the_marker() {
    assert_eq!(
        parse_one("\\begin{aligned}a &= b\\end{aligned}"),
        MathNode::EqArray {
            rows: vec![vec![
                text("a"),
                MathNode::Operator(Operator::Equals),
                text("b")
            ]],
            properties: None,
        }
    );
}

#[test]
fn parses_substack_in_both_spellings() {
    let expected = MathNode::EqArray {
        rows: vec![vec![text("a")], vec![text("b")]],
        properties: None,
    };
    assert_eq!(parse_one("\\substack{a \\\\ b}"), expected);
    assert_eq!(
        parse_one("\\begin{substack}a \\\\ b\\end{substack}"),
        expected
    );
}

#[test]
fn drops_the_empty_row_left_by_a_trailing_row_separator() {
    let MathNode::Matrix { rows, .. } = parse_one("\\begin{matrix}a \\\\ b \\\\ \\end{matrix}")
    else {
        panic!("expected a matrix");
    };
    assert_eq!(rows.len(), 2);
}

#[test]
fn nests_environments() {
    let MathNode::Matrix { rows, .. } =
        parse_one("\\begin{pmatrix}\\begin{matrix}a\\end{matrix}\\end{pmatrix}")
    else {
        panic!("expected a matrix");
    };
    assert!(matches!(rows[0][0].as_slice(), [MathNode::Matrix { .. }]));
}

// -- symbol vocabulary ---------------------------------------------------

#[test]
fn parses_greek_letters() {
    for (input, expected) in [
        ("\\alpha", PredefinedSymbol::Alpha),
        ("\\beta", PredefinedSymbol::Beta),
        ("\\pi", PredefinedSymbol::Pi),
        ("\\omega", PredefinedSymbol::Omega),
        ("\\Gamma", PredefinedSymbol::GammaCap),
        ("\\Omega", PredefinedSymbol::OmegaCap),
        ("\\infty", PredefinedSymbol::Infinity),
        ("\\aleph", PredefinedSymbol::Aleph),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::PredefinedSymbol(expected),
            "`{input}` should be a predefined symbol"
        );
    }
}

#[test]
fn parses_operator_commands() {
    for (input, expected) in [
        ("\\pm", Operator::PlusMinus),
        ("\\times", Operator::Times),
        ("\\cdot", Operator::Dot),
        ("\\leq", Operator::LessThanOrEqual),
        ("\\in", Operator::In),
        ("\\approx", Operator::Approx),
        ("\\partial", Operator::Partial),
        ("\\nabla", Operator::Nabla),
        ("\\rightarrow", Operator::RightArrow),
        ("\\forall", Operator::ForAll),
        ("\\cdots", Operator::CDots),
    ] {
        assert_eq!(
            parse_one(input),
            MathNode::Operator(expected),
            "`{input}` should be an operator"
        );
    }
}

#[test]
fn keeps_symbols_the_renderer_knows_by_name() {
    let MathNode::Symbol(symbol) = parse_one("\\hbar") else {
        panic!("expected a symbol");
    };
    assert_eq!(symbol.name.as_ref(), "hbar");
    assert_eq!(symbol.unicode, None);
}

#[test]
fn degrades_an_unknown_command_to_a_symbol() {
    let MathNode::Symbol(symbol) = parse_one("\\notarealcommand") else {
        panic!("expected a symbol");
    };
    assert_eq!(symbol.name.as_ref(), "notarealcommand");
}

#[test]
fn keeps_an_unpaired_delimiter_command_renderable() {
    let MathNode::Symbol(symbol) = parse_one("\\|") else {
        panic!("expected a symbol");
    };
    assert_eq!(symbol.unicode, Some('‖'));
}

#[test]
fn ignores_presentation_only_commands() {
    assert_eq!(parse("\\displaystyle x"), vec![text("x")]);
    assert_eq!(parse("\\textstyle x"), vec![text("x")]);
}

#[test]
fn drops_a_stray_alignment_marker_outside_an_environment() {
    assert_eq!(parse("a & b"), vec![text("a"), text("b")]);
}
