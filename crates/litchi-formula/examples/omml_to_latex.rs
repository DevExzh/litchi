//! Convert several OMML (Office Math Markup Language) snippets to LaTeX.
//!
//! Run with:
//!
//! ```bash
//! cargo run -p litchi-formula --example omml_to_latex --all-features
//! ```
//!
//! The example feeds a handful of representative formulas (a Pythagorean
//! identity, a fraction, an integral, and a square root) through
//! [`OmmlParser`] and then through [`LatexConverter`], printing both the
//! input markup and the resulting LaTeX for each one.

use litchi_formula::{Formula, LatexConverter, OmmlParser, omml_to_latex};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // A handful of representative OMML fragments.
    //
    // Each fragment is wrapped in a `<m:oMath>` root with the standard math
    // namespace declared, which is what the parser expects to see at the top
    // level of an OMML island inside an OOXML document.
    let samples: &[(&str, &str)] = &[
        (
            "x^2 + y^2 = z^2",
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:sSup>
                   <m:e><m:r><m:t>x</m:t></m:r></m:e>
                   <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                 </m:sSup>
                 <m:r><m:t>+</m:t></m:r>
                 <m:sSup>
                   <m:e><m:r><m:t>y</m:t></m:r></m:e>
                   <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                 </m:sSup>
                 <m:r><m:t>=</m:t></m:r>
                 <m:sSup>
                   <m:e><m:r><m:t>z</m:t></m:r></m:e>
                   <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                 </m:sSup>
               </m:oMath>"#,
        ),
        (
            "1/2",
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:f>
                   <m:num><m:r><m:t>1</m:t></m:r></m:num>
                   <m:den><m:r><m:t>2</m:t></m:r></m:den>
                 </m:f>
               </m:oMath>"#,
        ),
        (
            "definite integral of x dx from 0 to 1",
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:nary>
                   <m:naryPr><m:chr m:val="&#8747;"/></m:naryPr>
                   <m:sub><m:r><m:t>0</m:t></m:r></m:sub>
                   <m:sup><m:r><m:t>1</m:t></m:r></m:sup>
                   <m:e><m:r><m:t>x</m:t></m:r><m:r><m:t>dx</m:t></m:r></m:e>
                 </m:nary>
               </m:oMath>"#,
        ),
        (
            "square root of (a + b)",
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:rad>
                   <m:radPr><m:degHide m:val="1"/></m:radPr>
                   <m:deg/>
                   <m:e>
                     <m:r><m:t>a</m:t></m:r>
                     <m:r><m:t>+</m:t></m:r>
                     <m:r><m:t>b</m:t></m:r>
                   </m:e>
                 </m:rad>
               </m:oMath>"#,
        ),
    ];

    println!("== Using the high-level helper `omml_to_latex` ==\n");
    for (label, omml) in samples {
        println!("Formula : {label}");
        println!("OMML    : {}", oneline(omml));
        match omml_to_latex(omml) {
            Ok(latex) => println!("LaTeX   : {latex}"),
            Err(e) => println!("error   : {e}"),
        }
        println!();
    }

    // The same conversion done by hand, showing the building blocks the
    // helper composes internally. This is useful when you want to reuse a
    // single converter across many formulas to amortize allocations.
    println!("== Using `OmmlParser` and `LatexConverter` directly ==\n");
    let mut converter = LatexConverter::new();
    for (label, omml) in samples {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());
        let nodes = match parser.parse(omml) {
            Ok(nodes) => nodes,
            Err(e) => {
                println!("Formula : {label}");
                println!("error   : {e}");
                println!();
                continue;
            },
        };

        let mut formula = Formula::new();
        formula.set_root(nodes);

        let latex = converter.convert(&formula)?;
        println!("Formula : {label}");
        println!("LaTeX   : {latex}");
        println!();
    }

    Ok(())
}

/// Squash whitespace so OMML prints on a single line in the demo output.
fn oneline(s: &str) -> String {
    s.split_whitespace().collect::<Vec<_>>().join(" ")
}
