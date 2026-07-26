//! Example demonstrating RTF writing with new features:
//! - Headers and Footers
//! - Footnotes and Endnotes
//! - Hyperlinks
//! - Track Changes (Revisions)

use litchi::rtf::*;
use std::borrow::Cow;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating RTF document with new features using public writer API...\n");

    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);

    // Write document header
    writer.write_document_header()?;

    // Create and write a section with headers and footers
    println!("Adding section with headers and footers...");
    let mut section = Section::new();

    // Create header
    let mut header = HeaderFooter::new(HeaderFooterType::Header);
    header.add_paragraph(HeaderFooterParagraph::new(
        Cow::Borrowed("Document Header - Page "),
        Formatting::default(),
        Paragraph::default(),
    ));
    section.add_header_footer(header);

    // Create footer
    let mut footer = HeaderFooter::new(HeaderFooterType::Footer);
    footer.add_paragraph(HeaderFooterParagraph::new(
        Cow::Borrowed("Document Footer - © 2026"),
        Formatting::default(),
        Paragraph::default(),
    ));
    section.add_header_footer(footer);

    // Create first page header (different from regular header)
    let mut header_first = HeaderFooter::new(HeaderFooterType::HeaderFirst);
    header_first.add_paragraph(HeaderFooterParagraph::new(
        Cow::Borrowed("First Page Header"),
        Formatting {
            bold: true,
            font_size: std::num::NonZeroU16::new(28).unwrap(),
            ..Default::default()
        },
        Paragraph::default(),
    ));
    section.add_header_footer(header_first);

    writer.write_section(&section)?;

    // Write main document content
    println!("Adding main content with formatting...");
    writer.write_str("{\\b\\fs32 RTF Document with Advanced Features\\par}")?;
    writer.write_str("\\par")?;

    // Write paragraph with hyperlink
    println!("Adding hyperlink...");
    writer.write_str("{\\fs24 Visit our website: ")?;
    writer.write_hyperlink("https://github.com/DevExzh/litchi", "Litchi on GitHub")?;
    writer.write_str(" for more information.\\par}\\par")?;

    // Write paragraph with footnote
    println!("Adding footnote...");
    writer.write_str("{\\fs24 This is a paragraph with a footnote")?;
    let footnote = Note::footnote(
        Cow::Borrowed("1"),
        Cow::Borrowed("This is the footnote text providing additional information."),
    );
    writer.write_note(&footnote)?;
    writer.write_str(".\\par}\\par")?;

    // Write paragraph with endnote
    println!("Adding endnote...");
    writer.write_str("{\\fs24 This paragraph references an endnote")?;
    let endnote = Note::endnote(
        Cow::Borrowed("i"),
        Cow::Borrowed("This is an endnote with supplementary information."),
    );
    writer.write_note(&endnote)?;
    writer.write_str(".\\par}\\par")?;

    // Write paragraph with track changes (revision)
    println!("Adding track changes...");
    writer.write_str("{\\fs24 This document contains ")?;
    let revision = Revision::insertion(
        Cow::Borrowed("John Doe"),
        Cow::Borrowed("newly inserted text"),
    );
    writer.write_revision(&revision)?;
    writer.write_str(" as part of the review process.\\par}\\par")?;

    // Write another paragraph with deletion revision
    writer.write_str("{\\fs24 This paragraph shows ")?;
    let deletion = Revision::deletion(
        Cow::Borrowed("Jane Smith"),
        Cow::Borrowed("deleted content"),
    );
    writer.write_revision(&deletion)?;
    writer.write_str(" in the document.\\par}\\par")?;

    // Write a field example (PAGE field)
    println!("Adding field...");
    let page_field = Field::new(FieldType::Page, Cow::Borrowed("PAGE"), Cow::Borrowed("1"));
    writer.write_str("{\\fs24 Current page: ")?;
    writer.write_field(&page_field)?;
    writer.write_str("\\par}\\par")?;

    // Write multiple hyperlinks
    println!("Adding multiple hyperlinks...");
    writer.write_str("{\\fs24\\b Related Links:\\par}\\par")?;
    writer.write_str("{\\fs20 - Documentation: ")?;
    writer.write_hyperlink("https://docs.rs/litchi", "Rust Docs")?;
    writer.write_str("\\par}")?;
    writer.write_str("{\\fs20 - Repository: ")?;
    writer.write_hyperlink("https://github.com/DevExzh/litchi", "GitHub")?;
    writer.write_str("\\par}\\par")?;

    // Write complex footnote example
    println!("Adding complex footnote with formatting...");
    writer.write_str("{\\fs24 Advanced features")?;
    let mut complex_footnote = Note::footnote(
        Cow::Borrowed("2"),
        Cow::Borrowed(
            "This footnote includes detailed technical information about RTF 1.9.1 specification compliance.",
        ),
    );
    complex_footnote.formatting = Formatting {
        italic: true,
        font_size: std::num::NonZeroU16::new(16).unwrap(),
        ..Default::default()
    };
    writer.write_note(&complex_footnote)?;
    writer.write_str(" are fully supported.\\par}\\par")?;

    // Close document
    writer.write_str("}")?;
    writer.flush()?;

    // Save to file
    let filename = "rtf_new_features_output.rtf";
    std::fs::write(filename, &output)?;

    println!("\n✅ RTF document created successfully!");
    println!("📄 File saved as: {}", filename);
    println!("📊 File size: {} bytes", output.len());
    println!("\nFeatures included:");
    println!("  ✓ Headers and Footers (regular, first page, left/right)");
    println!("  ✓ Footnotes (with reference numbers)");
    println!("  ✓ Endnotes (with reference markers)");
    println!("  ✓ Hyperlinks (with URLs and display text)");
    println!("  ✓ Track Changes (insertions and deletions)");
    println!("  ✓ Fields (PAGE field example)");

    Ok(())
}
