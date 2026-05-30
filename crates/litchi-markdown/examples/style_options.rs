//! Compare `MarkdownOptions` style variants on a single input.
//!
//! Renders a tiny `MathSnippet` (a subscript fragment + a superscript fragment +
//! an inline formula) through every relevant variant of [`TableStyle`],
//! [`FormulaStyle`], and [`ScriptStyle`] so you can eyeball the differences.
//!
//! Note on naming: the task brief mentioned `TableStyle::Pipe` and
//! `ScriptStyle::Markdown`, but the real enum variants in `litchi-markdown` are
//! `TableStyle::{Markdown, MinimalHtml, StyledHtml}` and
//! `ScriptStyle::{Html, Unicode}`. This example uses the actual variants.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-markdown --example style_options
//! ```
use std::fmt::Write as _;

use litchi_core::Result;
use litchi_markdown::{
    FormulaStyle, MarkdownOptions, ScriptStyle, StrikethroughStyle, TableStyle, ToMarkdown,
    unicode::{convert_to_subscript, convert_to_superscript},
};

/// A tiny "math snippet" with a subscript label, a superscript exponent, and an
/// inline formula. The struct itself stores raw text; rendering decisions are
/// driven entirely by [`MarkdownOptions`].
struct MathSnippet {
    /// Raw text to render in subscript position, e.g. `"i+1"`.
    subscript: String,
    /// Raw text to render in superscript position, e.g. `"n2"`.
    superscript: String,
    /// LaTeX source for an inline formula, without delimiters, e.g. `"a^2+b^2=c^2"`.
    formula: String,
    /// One-row, one-cell table used purely to show `TableStyle` differences.
    cell: String,
}

impl ToMarkdown for MathSnippet {
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut out = String::with_capacity(256);

        // --- Script rendering -------------------------------------------------
        // ScriptStyle::Html -> wrap in <sub>/<sup>; preserves all characters.
        // ScriptStyle::Unicode -> map char-by-char via `unicode` helpers, falling
        // back to original chars where no Unicode equivalent exists.
        match options.script_style {
            ScriptStyle::Html => {
                writeln!(out, "x<sub>{}</sub>", self.subscript).unwrap();
                writeln!(out, "x<sup>{}</sup>", self.superscript).unwrap();
            }
            ScriptStyle::Unicode => {
                writeln!(out, "x{}", convert_to_subscript(&self.subscript)).unwrap();
                writeln!(out, "x{}", convert_to_superscript(&self.superscript)).unwrap();
            }
        }

        // --- Formula rendering ------------------------------------------------
        match options.formula_style {
            FormulaStyle::LaTeX => writeln!(out, "Inline: \\({}\\)", self.formula).unwrap(),
            FormulaStyle::Dollar => writeln!(out, "Inline: ${}$", self.formula).unwrap(),
        }

        // --- Strikethrough sample (so the option isn't silent) ---------------
        match options.strikethrough_style {
            StrikethroughStyle::Markdown => writeln!(out, "~~old~~").unwrap(),
            StrikethroughStyle::Html => writeln!(out, "<del>old</del>").unwrap(),
        }

        // --- Single-cell table ----------------------------------------------
        match options.table_style {
            TableStyle::Markdown => {
                writeln!(out, "| Header |").unwrap();
                writeln!(out, "|--------|").unwrap();
                writeln!(out, "| {} |", self.cell).unwrap();
            }
            TableStyle::MinimalHtml => {
                writeln!(
                    out,
                    "<table><tr><th>Header</th></tr><tr><td>{}</td></tr></table>",
                    self.cell
                )
                .unwrap();
            }
            TableStyle::StyledHtml => {
                let pad = " ".repeat(options.html_table_indent);
                writeln!(out, "<table>").unwrap();
                writeln!(out, "{pad}<tr>").unwrap();
                writeln!(out, "{pad}{pad}<th>Header</th>").unwrap();
                writeln!(out, "{pad}</tr>").unwrap();
                writeln!(out, "{pad}<tr>").unwrap();
                writeln!(out, "{pad}{pad}<td>{}</td>", self.cell).unwrap();
                writeln!(out, "{pad}</tr>").unwrap();
                writeln!(out, "</table>").unwrap();
            }
        }

        Ok(out)
    }
}

fn render(label: &str, snippet: &MathSnippet, options: &MarkdownOptions) -> Result<()> {
    println!("===== {label} =====");
    println!("{}", snippet.to_markdown_with_options(options)?);
    Ok(())
}

fn main() -> Result<()> {
    let snippet = MathSnippet {
        subscript: "i+1".to_owned(),
        superscript: "n2".to_owned(),
        formula: "a^2+b^2=c^2".to_owned(),
        cell: "value".to_owned(),
    };

    // TableStyle variants (Markdown, MinimalHtml, StyledHtml).
    render(
        "TableStyle::Markdown (default)",
        &snippet,
        &MarkdownOptions::new().with_table_style(TableStyle::Markdown),
    )?;
    render(
        "TableStyle::MinimalHtml",
        &snippet,
        &MarkdownOptions::new().with_table_style(TableStyle::MinimalHtml),
    )?;
    render(
        "TableStyle::StyledHtml (indent=2)",
        &snippet,
        &MarkdownOptions::new()
            .with_table_style(TableStyle::StyledHtml)
            .with_html_table_indent(2),
    )?;

    // FormulaStyle variants.
    render(
        "FormulaStyle::LaTeX (default)",
        &snippet,
        &MarkdownOptions::new().with_formula_style(FormulaStyle::LaTeX),
    )?;
    render(
        "FormulaStyle::Dollar",
        &snippet,
        &MarkdownOptions::new().with_formula_style(FormulaStyle::Dollar),
    )?;

    // ScriptStyle variants -- this exercises the `unicode` helpers.
    render(
        "ScriptStyle::Html (default)",
        &snippet,
        &MarkdownOptions::new().with_script_style(ScriptStyle::Html),
    )?;
    render(
        "ScriptStyle::Unicode",
        &snippet,
        &MarkdownOptions::new().with_script_style(ScriptStyle::Unicode),
    )?;

    // StrikethroughStyle variants, for completeness.
    render(
        "StrikethroughStyle::Markdown (default)",
        &snippet,
        &MarkdownOptions::new().with_strikethrough_style(StrikethroughStyle::Markdown),
    )?;
    render(
        "StrikethroughStyle::Html",
        &snippet,
        &MarkdownOptions::new().with_strikethrough_style(StrikethroughStyle::Html),
    )?;

    Ok(())
}
