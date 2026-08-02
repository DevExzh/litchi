//! Paragraph formatting DOC writer example
//!
//! Demonstrates paragraph alignment, indents, and spacing.
//!
//! Run with:
//!   cargo run --example doc_paragraph_format --features doc --no-default-features -- <output.doc>

use litchi::doc::{CharacterFormatting, DocWriter, ParagraphFormatting};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let out = env::args()
        .nth(1)
        .unwrap_or_else(|| "doc_paragraph_format.doc".to_string());

    let mut doc = DocWriter::new();

    // Left aligned with indents
    let p1 = ParagraphFormatting {
        alignment: Some(0),           // left
        left_indent: Some(720),       // 0.5 inch
        first_line_indent: Some(360), // 0.25 inch
        ..Default::default()
    };
    doc.add_paragraph_with_format(
        "Left aligned with indents",
        CharacterFormatting::default(),
        p1,
    )?;

    // Centered with spacing before/after
    let p2 = ParagraphFormatting {
        alignment: Some(1),      // center
        space_before: Some(240), // 12pt
        space_after: Some(240),  // 12pt
        ..Default::default()
    };
    doc.add_paragraph_with_format("Centered with spacing", CharacterFormatting::default(), p2)?;

    // Justified with right indent
    let p3 = ParagraphFormatting {
        alignment: Some(3),      // justify
        right_indent: Some(720), // 0.5 inch
        ..Default::default()
    };
    doc.add_paragraph_with_format(
        "Justified with right indent",
        CharacterFormatting::default(),
        p3,
    )?;

    doc.save(&out)?;
    println!("Paragraph formatting DOC written to: {}", out);
    Ok(())
}
