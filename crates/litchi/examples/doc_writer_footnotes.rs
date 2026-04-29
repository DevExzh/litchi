//! Example demonstrating footnotes and endnotes in DOC files
//!
//! This example creates a DOC file with:
//! - Multiple footnotes with text stored in the footnote subdocument
//! - Multiple endnotes with text stored in the endnote subdocument
//! - Reference positions in the main document body
//!
//! Run with: cargo run --example doc_writer_footnotes

use litchi::ole::doc::writer::{DocWriter, FootnoteEntry};
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOC file with footnotes and endnotes...");

    let mut writer = DocWriter::new();

    // Add main body content
    writer.add_paragraph("Understanding Footnotes and Endnotes")?;
    writer.add_paragraph("")?;

    // Paragraph 1 — footnotes reference positions are character offsets in the main text
    writer.add_paragraph(
        "Footnotes appear at the bottom of the page. They provide additional information \
         or citations without cluttering the main text. Footnotes are commonly used in \
         academic writing.",
    )?;
    writer.add_paragraph("")?;

    // Paragraph 2 — endnotes
    writer.add_paragraph(
        "Endnotes, on the other hand, appear at the end of the document. They serve a \
         similar purpose to footnotes but are collected in one place. Some authors prefer \
         endnotes to maintain page flow.",
    )?;
    writer.add_paragraph("")?;

    writer.add_paragraph(
        "Both footnotes and endnotes are essential tools for scholarly writing. They allow \
         authors to provide references, explanations, or additional context.",
    )?;

    // Add footnotes — ref_position is the CP in the main document where the note reference sits
    writer.add_footnote(FootnoteEntry::new(
        50,
        "This is the first footnote, providing additional context.".to_string(),
        1,
    ));
    writer.add_footnote(FootnoteEntry::new(
        120,
        "Second footnote: See Smith, J. (2023). 'The Art of Annotation'.".to_string(),
        2,
    ));
    writer.add_footnote(FootnoteEntry::new(
        200,
        "Third footnote: Footnotes have been used since medieval manuscripts.".to_string(),
        3,
    ));

    // Add endnotes
    writer.add_endnote(FootnoteEntry::new(
        300,
        "Endnote 1: Endnotes are typically collected at the document end.".to_string(),
        1,
    ));
    writer.add_endnote(FootnoteEntry::new(
        400,
        "Endnote 2: The choice between footnotes and endnotes is stylistic.".to_string(),
        2,
    ));

    println!("Footnotes configured: 3 entries");
    println!("Endnotes configured: 2 entries");

    // Save the document
    let output_path = "output/doc_footnotes.doc";
    writer.save(output_path)?;

    println!("\nDocument saved to: {}", output_path);
    println!("\nOpen this file in Microsoft Word to verify:");
    println!("   - View > References > Show Notes to see footnotes/endnotes");
    println!("   - Footnote text is stored in the footnote subdocument");
    println!("   - Endnote text is stored in the endnote subdocument");

    Ok(())
}
