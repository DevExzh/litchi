//! Integration tests for the LaTeX parser and its round trip through
//! [`LatexConverter`].
//!
//! Two properties are checked for every construct the parser supports:
//!
//! 1. **Rendering** — parsing an expression and converting the resulting AST
//!    back to LaTeX yields the expected string.
//! 2. **Idempotence** — feeding that string back through the parser and the
//!    converter reproduces it exactly, so LaTeX -> AST -> LaTeX is a fixpoint
//!    and no information is lost on repeated conversions.

use litchi_formula::latex::{LatexConverter, LatexParseError, LatexParser};
use litchi_formula::{Formula, MathNode};

/// Parse `input` and convert the resulting AST back to LaTeX.
fn render(input: &str) -> String {
    let nodes = LatexParser::new()
        .parse(input)
        .unwrap_or_else(|error| panic!("`{input}` should parse but failed: {error}"));
    let mut converter = LatexConverter::new();
    converter
        .convert_nodes(&nodes)
        .unwrap_or_else(|error| panic!("`{input}` should convert but failed: {error}"))
        .to_string()
}

/// Every supported construct paired with its rendering.
///
/// Differences from the input are intentional normalisations: `\dfrac` and
/// `\tfrac` collapse onto `\frac`, `\widehat` onto `\hat`, and matrix
/// environments onto the spelling `MatrixFence` renders.
const ROUND_TRIP_CASES: &[(&str, &str)] = &[
    ("x", "x"),
    ("abc", "abc"),
    ("1234", "1234"),
    ("3.14", "3.14"),
    ("x_i", "x_{i}"),
    ("x^2", "x^{2}"),
    ("x_i^2", "x_{i}^{2}"),
    ("x^{a+b}", "x^{a+b}"),
    ("{a+b}^2", "{a+b}^{2}"),
    ("\\frac{a}{b}", "\\frac{a}{b}"),
    ("\\dfrac{a}{b}", "\\frac{a}{b}"),
    ("\\tfrac{a}{b}", "\\frac{a}{b}"),
    ("{a \\over b}", "\\frac{a}{b}"),
    ("\\binom{n}{k}", "\\left(\\frac{n}{k}\\right)"),
    ("{n \\choose k}", "\\left(\\frac{n}{k}\\right)"),
    ("\\sqrt{x}", "\\sqrt{x}"),
    ("\\sqrt[3]{x}", "\\sqrt[3]{x}"),
    ("\\left( x \\right)", "\\left(x\\right)"),
    ("\\left[ x \\right]", "\\left[x\\right]"),
    ("\\left\\{ x \\right\\}", "\\left\\{x\\right\\}"),
    ("\\left| x \\right|", "\\left|x\\right|"),
    ("\\left\\| x \\right\\|", "\\left\\|x\\right\\|"),
    (
        "\\left\\langle x \\right\\rangle",
        "\\left\\langle x\\right\\rangle",
    ),
    (
        "\\left\\lfloor x \\right\\rfloor",
        "\\left\\lfloor x\\right\\rfloor",
    ),
    (
        "\\left\\lceil x \\right\\rceil",
        "\\left\\lceil x\\right\\rceil",
    ),
    ("\\left. x \\right)", "\\left.x\\right)"),
    ("(a+b)", "\\left(a+b\\right)"),
    ("[a]", "\\left[a\\right]"),
    ("\\sum_{i=1}^{n} a", "\\sum_{i=1}^{n}a"),
    ("\\prod_{i}", "\\prod_{i}"),
    ("\\coprod", "\\coprod"),
    ("\\int_0^1", "\\int_{0}^{1}"),
    ("\\iint", "\\iint"),
    ("\\iiint", "\\iiint"),
    ("\\oint", "\\oint"),
    ("\\bigcup_{i}", "\\bigcup_{i}"),
    ("\\bigcap", "\\bigcap"),
    ("\\lim_{x \\to 0}", "\\lim_{x\\to 0}"),
    ("\\sum\\limits_{i}", "\\sum_{i}"),
    ("\\sum\\nolimits_{i}", "\\sum_{i}"),
    ("\\sin x", "\\sin{x}"),
    ("\\cos x", "\\cos{x}"),
    ("\\tan x", "\\tan{x}"),
    ("\\sec x", "\\sec{x}"),
    ("\\csc x", "\\csc{x}"),
    ("\\cot x", "\\cot{x}"),
    ("\\arcsin x", "\\arcsin{x}"),
    ("\\arccos x", "\\arccos{x}"),
    ("\\arctan x", "\\arctan{x}"),
    ("\\sinh x", "\\sinh{x}"),
    ("\\cosh x", "\\cosh{x}"),
    ("\\tanh x", "\\tanh{x}"),
    ("\\log x", "\\log{x}"),
    ("\\ln x", "\\ln{x}"),
    ("\\exp x", "\\exp{x}"),
    ("\\min x", "\\min{x}"),
    ("\\max x", "\\max{x}"),
    ("\\sup x", "\\sup{x}"),
    ("\\inf x", "\\inf{x}"),
    ("\\det x", "\\det{x}"),
    ("\\dim x", "\\dim{x}"),
    ("\\ker x", "\\ker{x}"),
    ("\\deg x", "\\deg{x}"),
    ("\\gcd x", "\\gcd{x}"),
    ("\\arg x", "\\arg{x}"),
    ("\\Pr x", "\\operatorname{Pr}{x}"),
    ("\\bmod x", "\\mod{x}"),
    ("\\hat{x}", "\\hat{x}"),
    ("\\widehat{x}", "\\hat{x}"),
    ("\\tilde{x}", "\\tilde{x}"),
    ("\\widetilde{x}", "\\tilde{x}"),
    ("\\bar{x}", "\\bar{x}"),
    ("\\vec{x}", "\\vec{x}"),
    ("\\dot{x}", "\\dot{x}"),
    ("\\ddot{x}", "\\ddot{x}"),
    ("\\dddot{x}", "\\dddot{x}"),
    ("\\acute{x}", "\\acute{x}"),
    ("\\grave{x}", "\\grave{x}"),
    ("\\check{x}", "\\check{x}"),
    ("\\breve{x}", "\\breve{x}"),
    ("\\overline{x}", "\\bar{x}"),
    ("\\underline{x}", "\\underset{}{x}"),
    ("\\overbrace{x}", "\\overbrace{x}"),
    ("\\underbrace{x}", "\\underbrace{x}"),
    ("\\mathrm{x}", "\\mathrm{x}"),
    ("\\mathbf{x}", "\\mathbf{x}"),
    ("\\mathit{x}", "\\mathit{x}"),
    ("\\mathbb{x}", "\\mathbb{x}"),
    ("\\mathcal{x}", "\\mathcal{x}"),
    ("\\mathfrak{x}", "\\mathfrak{x}"),
    ("\\mathsf{x}", "\\mathsf{x}"),
    ("\\mathtt{x}", "\\mathtt{x}"),
    ("\\boldsymbol{x}", "\\bm{x}"),
    ("\\text{a b}", "\\text{a\\ b}"),
    ("\\textbf{a b}", "\\mathbf{\\text{a\\ b}}"),
    ("\\textit{a b}", "\\mathit{\\text{a\\ b}}"),
    ("\\phantom{x}", "\\phantom{x}"),
    ("\\boxed{x}", "\\boxed{x}"),
    ("\\fbox{x}", "\\boxed{x}"),
    ("a\\,b", "a\\,b"),
    ("a\\:b", "a\\:b"),
    ("a\\;b", "a\\;b"),
    ("a\\!b", "a\\!b"),
    ("a\\quad b", "a\\quad b"),
    ("a\\qquad b", "a\\qquad b"),
    ("a\\ b", "a\\:b"),
    ("a\\\\b", "a\\\\b"),
    (
        "\\begin{matrix}a & b \\\\ c & d\\end{matrix}",
        "\\begin{matrix}a & b \\\\ c & d\\end{matrix}",
    ),
    (
        "\\begin{pmatrix}a\\end{pmatrix}",
        "\\begin{pmatrix}a\\end{pmatrix}",
    ),
    (
        "\\begin{bmatrix}a\\end{bmatrix}",
        "\\begin{bmatrix}a\\end{bmatrix}",
    ),
    (
        "\\begin{Bmatrix}a\\end{Bmatrix}",
        "\\begin{Bmatrix}a\\end{Bmatrix}",
    ),
    (
        "\\begin{vmatrix}a\\end{vmatrix}",
        "\\begin{vmatrix}a\\end{vmatrix}",
    ),
    (
        "\\begin{Vmatrix}a\\end{Vmatrix}",
        "\\begin{Vmatrix}a\\end{Vmatrix}",
    ),
    (
        "\\begin{smallmatrix}a\\end{smallmatrix}",
        "\\begin{matrix}a\\end{matrix}",
    ),
    (
        "\\begin{array}{lc}a & b\\end{array}",
        "\\begin{matrix}a & b\\end{matrix}",
    ),
    (
        "\\begin{cases}a & b\\end{cases}",
        "\\begin{Bmatrix}a & b\\end{Bmatrix}",
    ),
    (
        "\\begin{aligned}a \\\\ b\\end{aligned}",
        "\\begin{align*}a\\\\b\\end{align*}",
    ),
    (
        "\\begin{align}a\\end{align}",
        "\\begin{align*}a\\end{align*}",
    ),
    (
        "\\begin{align*}a\\end{align*}",
        "\\begin{align*}a\\end{align*}",
    ),
    (
        "\\begin{gathered}a\\end{gathered}",
        "\\begin{align*}a\\end{align*}",
    ),
    (
        "\\begin{gather}a\\end{gather}",
        "\\begin{align*}a\\end{align*}",
    ),
    (
        "\\begin{split}a\\end{split}",
        "\\begin{align*}a\\end{align*}",
    ),
    ("\\substack{a \\\\ b}", "\\begin{align*}a\\\\b\\end{align*}"),
    ("\\alpha\\beta\\gamma x", "\\alpha\\beta\\gamma x"),
    ("\\Gamma\\Delta", "\\Gamma\\Delta"),
    ("\\infty", "\\infty"),
    ("\\aleph", "\\aleph"),
    ("\\hbar", "\\hbar"),
    ("a+b-c=d<e>f", "a+b-c=d<e>f"),
    ("a!b,c*d/e", "a!b,c*d/e"),
    ("α+β", "\\alpha+\\beta"),
    ("x'", "x'"),
    ("{}_a^b X", "\\presubsup{X}{a}{b}"),
    ("\\underset{u}{b}", "\\underset{u}{b}"),
    ("\\overset{o}{b}", "\\overset{o}{b}"),
    ("\\operatorname{foo}{x}", "\\operatorname{foo}{x}"),
];

#[test]
fn every_construct_renders_as_expected() {
    for (input, expected) in ROUND_TRIP_CASES {
        assert_eq!(&render(input), expected, "rendering of `{input}`");
    }
}

#[test]
fn latex_to_ast_to_latex_is_a_fixpoint() {
    for (input, expected) in ROUND_TRIP_CASES {
        let once = render(input);
        let twice = render(&once);
        assert_eq!(once, twice, "`{input}` is not stable under re-parsing");
        assert_eq!(&once, expected);
    }
}

#[test]
fn parsed_nodes_borrow_from_the_input() {
    let input = String::from("\\text{a b} 42 xy");
    let nodes = LatexParser::new().parse(&input).expect("parses");

    let mut borrowed = 0usize;
    fn count(nodes: &[MathNode<'_>], borrowed: &mut usize) {
        for node in nodes {
            match node {
                MathNode::Text(value) | MathNode::Number(value) => {
                    assert!(
                        matches!(value, std::borrow::Cow::Borrowed(_)),
                        "{value:?} should borrow from the input"
                    );
                    *borrowed += 1;
                },
                MathNode::Run { content, .. } => count(content, borrowed),
                _ => {},
            }
        }
    }
    count(&nodes, &mut borrowed);
    assert!(borrowed >= 3, "expected several borrowed literals");
}

#[test]
fn a_parsed_expression_can_populate_a_formula() {
    let nodes = LatexParser::new()
        .parse("\\frac{a}{b}")
        .expect("expression parses");

    let mut formula = Formula::new();
    formula.set_root(nodes);

    let mut converter = LatexConverter::new();
    let latex = converter.convert(&formula).expect("formula converts");
    assert_eq!(latex, "\\[\\frac{a}{b}\\]");
}

#[test]
fn converter_output_for_a_whole_formula_parses_back() {
    let mut formula = Formula::new();
    formula.set_root(
        LatexParser::new()
            .parse("\\sqrt[3]{x} + \\alpha")
            .expect("expression parses"),
    );

    let mut converter = LatexConverter::new();
    let latex = converter.convert(&formula).expect("converts").to_string();
    // The display-math delimiters `\[ ... \]` are accepted and stripped.
    assert_eq!(render(&latex), "\\sqrt[3]{x}+\\alpha");
}

#[test]
fn unknown_commands_degrade_instead_of_failing() {
    let nodes = LatexParser::new()
        .parse("\\notarealcommand{x}")
        .expect("unknown commands do not fail the parse");
    assert!(matches!(nodes.first(), Some(MathNode::Symbol(_))));
}

/// Predicate identifying the error variant a malformed input should produce.
type ErrorCheck = fn(&LatexParseError) -> bool;

#[test]
fn structurally_broken_input_is_reported_not_ignored() {
    let cases: &[(&str, ErrorCheck)] = &[
        ("{a", |error| {
            matches!(error, LatexParseError::UnmatchedGroupOpen { .. })
        }),
        ("a}", |error| {
            matches!(error, LatexParseError::UnmatchedGroupClose { .. })
        }),
        ("\\left(x", |error| {
            matches!(error, LatexParseError::UnmatchedLeft { .. })
        }),
        ("x\\right)", |error| {
            matches!(error, LatexParseError::UnmatchedRight { .. })
        }),
        ("\\begin{matrix}x", |error| {
            matches!(error, LatexParseError::UnclosedEnvironment { .. })
        }),
        ("\\begin{matrix}x\\end{cases}", |error| {
            matches!(error, LatexParseError::MismatchedEnvironment { .. })
        }),
        ("x^1^2", |error| {
            matches!(error, LatexParseError::DuplicateScript { .. })
        }),
        ("\\frac{a}", |error| {
            matches!(error, LatexParseError::MissingArgument { .. })
        }),
        ("x \\", |error| {
            matches!(error, LatexParseError::IncompleteCommand { .. })
        }),
    ];

    for (input, matches_expected) in cases {
        let error = LatexParser::new()
            .parse(input)
            .expect_err(&format!("`{input}` should be rejected"));
        assert!(
            matches_expected(&error),
            "`{input}` produced an unexpected error: {error}"
        );
        assert!(!error.to_string().is_empty());
    }
}

#[test]
fn deeply_nested_input_is_rejected_rather_than_overflowing_the_stack() {
    let depth = 10_000;
    let deep = "{".repeat(depth) + &"}".repeat(depth);
    assert!(matches!(
        LatexParser::new().parse(&deep),
        Err(LatexParseError::NestingTooDeep { .. })
    ));

    let deep_fences = "\\left(".repeat(depth) + &"\\right)".repeat(depth);
    assert!(LatexParser::new().parse(&deep_fences).is_err());
}

#[test]
fn adversarial_input_never_panics() {
    let fragments = [
        "\\frac",
        "\\sqrt",
        "\\left",
        "\\right",
        "\\begin",
        "\\end",
        "{",
        "}",
        "_",
        "^",
        "&",
        "\\\\",
        "(",
        ")",
        "[",
        "]",
        "|",
        "\\{",
        "\\}",
        "a",
        "1",
        "+",
        "\\alpha",
        "\\sum",
        "\\text",
        "\\over",
        "\\choose",
        "\\limits",
        "%",
        "\n",
        " ",
        "$",
        "\\lim",
        "matrix",
        "*",
        "\\operatorname",
        "\\substack",
        "α",
        "\\|",
        ".",
        "'",
    ];

    // A deterministic xorshift keeps the sweep reproducible across runs.
    let mut state: u64 = 0x2545_F491_4F6C_DD1D;
    let mut next = move || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    };

    for _ in 0..20_000 {
        let len = (next() % 24) as usize;
        let mut input = String::new();
        for _ in 0..len {
            input.push_str(fragments[(next() as usize) % fragments.len()]);
        }
        if let Ok(nodes) = LatexParser::new().parse(&input) {
            let mut converter = LatexConverter::new();
            converter
                .convert_nodes(&nodes)
                .expect("conversion succeeds");
        }
    }
}
