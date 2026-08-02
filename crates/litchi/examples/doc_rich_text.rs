//! Rich text DOC writer example
//!
//! Demonstrates multiple runs with character formatting and a custom font table.
//!
//! Run with:
//!   cargo run --example doc_rich_text --features doc --no-default-features -- <output.doc>

use litchi::doc::{CharacterFormatting, DocWriter, ParagraphFormatting};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let out = env::args()
        .nth(1)
        .unwrap_or_else(|| "doc_rich_text.doc".to_string());

    let mut doc = DocWriter::new();

    // Paragraph 1: rich runs
    let runs = vec![
        ("Hello, ".to_string(), CharacterFormatting::default()),
        (
            "bold".to_string(),
            CharacterFormatting {
                bold: Some(true),
                ..Default::default()
            },
        ),
        (", ".to_string(), CharacterFormatting::default()),
        (
            "italic".to_string(),
            CharacterFormatting {
                italic: Some(true),
                ..Default::default()
            },
        ),
        (", ".to_string(), CharacterFormatting::default()),
        (
            "underlined".to_string(),
            CharacterFormatting {
                underline: Some(true),
                ..Default::default()
            },
        ),
        (
            ", size14 ".to_string(),
            CharacterFormatting {
                font_size: Some(28),
                ..Default::default()
            },
        ),
        (
            "Arial".to_string(),
            CharacterFormatting {
                font_name: Some("Arial".to_string()),
                ..Default::default()
            },
        ),
        (
            ", red".to_string(),
            CharacterFormatting {
                color: Some((255, 0, 0)),
                ..Default::default()
            },
        ),
    ];
    doc.add_paragraph_runs(runs, ParagraphFormatting::default())?;

    // Paragraph 2: plain text
    doc.add_paragraph("This is a second paragraph.")?;

    doc.save(&out)?;
    println!("Rich text DOC written to: {}", out);
    Ok(())
}
