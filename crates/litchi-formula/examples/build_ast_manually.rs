//! Build a small formula AST by hand and convert it to LaTeX.
//!
//! Run with:
//!
//! ```bash
//! cargo run -p litchi-formula --example build_ast_manually --all-features
//! ```
//!
//! This example skips the OMML/MTEF parsers entirely. It uses
//! [`Formula::new`] together with [`FormulaBuilder`] (and the [`MathNode`]
//! enum directly) to construct a small AST in code, then converts it to
//! LaTeX with [`LatexConverter`].
//!
//! The two formulas built are:
//!   1. `x^2 + 1` (a simple superscript followed by an addition and a
//!      number literal).
//!   2. `(a + b) / 2` (a fraction whose numerator is itself a small
//!      sub-expression).

use litchi_formula::{Formula, FormulaBuilder, LatexConverter, MathNode, Operator};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut converter = LatexConverter::new();

    // ---------------------------------------------------------------
    // Formula 1: x^2 + 1
    // ---------------------------------------------------------------
    //
    // We use two `Formula` values: the first owns the arena that backs
    // the AST nodes, and the second is what we hand to `LatexConverter`.
    // This split mirrors what `omml_to_latex()` does internally and is
    // needed because `FormulaBuilder` holds a shared borrow of the arena
    // for the whole construction phase, which prevents calling
    // `set_root` (a mutable borrow) on the same `Formula`.
    let arena_owner = Formula::new();
    let nodes = {
        let builder = FormulaBuilder::new(arena_owner.arena());

        // x^2 -- a Power node whose base is `x` and exponent is `2`.
        let power = builder.power(vec![builder.text("x")], vec![builder.number("2")]);

        // The trailing "+ 1" is rendered as an Operator followed by a
        // number literal. Using `MathNode::Operator` (rather than raw
        // `Text`) lets the LaTeX backend pick the right spacing.
        let plus = MathNode::Operator(Operator::Plus);
        let one = builder.number("1");

        vec![power, plus, one]
    };
    let mut formula = Formula::new();
    formula.set_root(nodes);
    println!("Formula 1 : x^2 + 1");
    println!("LaTeX     : {}", converter.convert(&formula)?);
    println!();

    // ---------------------------------------------------------------
    // Formula 2: (a + b) / 2
    // ---------------------------------------------------------------
    let arena_owner = Formula::new();
    let nodes = {
        let builder = FormulaBuilder::new(arena_owner.arena());

        // Numerator: `a + b`, written as three nodes that flow inline.
        let numerator: Vec<MathNode> = vec![
            builder.text("a"),
            MathNode::Operator(Operator::Plus),
            builder.text("b"),
        ];

        // Denominator: just a single number.
        let denominator: Vec<MathNode> = vec![builder.number("2")];

        vec![builder.frac(numerator, denominator)]
    };
    let mut formula = Formula::new();
    formula.set_root(nodes);
    println!("Formula 2 : (a + b) / 2");
    println!("LaTeX     : {}", converter.convert(&formula)?);

    // Keep `arena_owner` values alive until the end so the nodes they
    // back outlive the converter's reads. (In practice they are dropped
    // here at the end of `main`.)
    drop(arena_owner);

    Ok(())
}
