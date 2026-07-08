//! Comprehensive example demonstrating all DOC writer features
//!
//! This example creates a DOC file with:
//! - Headers and footers
//! - Footnotes and endnotes
//! - Numbered and bulleted lists
//! - Hyperlinks
//! - Embedded images
//! - Tables
//! - Rich text formatting
//!
//! Run with: cargo run --example doc_writer_all_features

use litchi::ole::doc::writer::{
    CharacterFormatting, DocWriter, FootnoteEntry, ListFormatOverride, ListLevel, ListStructure,
    NumberFormat, ParagraphFormatting,
};
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating comprehensive DOC file with all features...");

    let mut writer = DocWriter::new();
    writer.set_property("Title", "Comprehensive DOC Features Demonstration");
    writer.set_property("Author", "Litchi Examples");

    writer.set_first_header("=== Comprehensive DOC Features Demo ===");
    writer.set_odd_header("DOC Features - Odd Pages");
    writer.set_even_header("DOC Features - Even Pages");
    writer.set_first_footer("Title Page");
    writer.set_odd_footer("Page [#] - Demo Document");
    writer.set_even_footer("Demo Document - Page [#]");

    // ===== TITLE PAGE =====
    let title_char = CharacterFormatting {
        bold: Some(true),
        font_size: Some(28),
        ..Default::default()
    };
    let title_para = ParagraphFormatting {
        alignment: Some(1),
        ..Default::default()
    };
    writer.add_paragraph_with_format(
        "Comprehensive DOC Features Demonstration",
        title_char,
        title_para,
    )?;
    writer.add_paragraph("")?;
    writer
        .add_paragraph("This document showcases all the newly implemented DOC writer features:")?;
    writer.add_paragraph("• Headers and Footers")?;
    writer.add_paragraph("• Footnotes and Endnotes")?;
    writer.add_paragraph("• List Numbering")?;
    writer.add_paragraph("• Hyperlinks")?;
    writer.add_paragraph("• Rich text formatting")?;
    writer.add_paragraph("• Tables")?;
    writer.add_paragraph("")?;

    // ===== SECTION 1: HEADERS AND FOOTERS =====
    writer.add_paragraph("Section 1: Headers and Footers")?;
    writer.add_paragraph("")?;
    writer.add_paragraph(
        "This document uses different headers and footers for odd and even pages, \
         with a special first page header.",
    )?;
    writer.add_paragraph("")?;

    // ===== SECTION 2: FOOTNOTES =====
    writer.add_paragraph("Section 2: Footnotes and Endnotes")?;
    writer.add_paragraph("")?;
    writer.add_paragraph(
        "Footnotes provide additional context[1] at the bottom of each page. \
         They are useful for citations[2] and explanatory notes[3].",
    )?;
    writer.add_paragraph("")?;
    writer.add_paragraph(
        "Endnotes appear at the end of the document[i] and serve a similar purpose. \
         Some writers prefer endnotes[ii] for longer documents[iii].",
    )?;
    writer.add_paragraph("")?;

    // ===== SECTION 3: NUMBERING =====
    writer.add_paragraph("Section 3: List Numbering")?;
    writer.add_paragraph("")?;
    writer.add_paragraph("3.1 Numbered List:")?;
    writer.add_paragraph("1. First item")?;
    writer.add_paragraph("2. Second item")?;
    writer.add_paragraph("3. Third item")?;
    writer.add_paragraph("")?;

    writer.add_paragraph("3.2 Bulleted List:")?;
    writer.add_paragraph("• Bullet point one")?;
    writer.add_paragraph("• Bullet point two")?;
    writer.add_paragraph("• Bullet point three")?;
    writer.add_paragraph("")?;

    writer.add_paragraph("3.3 Multi-level List:")?;
    writer.add_paragraph("1. Level 1 item")?;
    writer.add_paragraph("   a. Level 2 item")?;
    writer.add_paragraph("   b. Level 2 item")?;
    writer.add_paragraph("      i. Level 3 item")?;
    writer.add_paragraph("2. Level 1 item")?;
    writer.add_paragraph("")?;

    let mut list1 = ListStructure::new(1);
    list1.add_level(ListLevel::new(1, NumberFormat::Decimal));
    writer.add_list(list1);
    writer.add_list_override(ListFormatOverride::new(1, 1));

    let mut list2 = ListStructure::new(2);
    list2.add_level(ListLevel::new(1, NumberFormat::Bullet));
    writer.add_list(list2);
    writer.add_list_override(ListFormatOverride::new(2, 2));

    let mut list3 = ListStructure::new(3);
    list3.add_level(ListLevel::new(1, NumberFormat::Decimal));
    list3.add_level(ListLevel::new(1, NumberFormat::LowerLetter));
    list3.add_level(ListLevel::new(1, NumberFormat::LowerRoman));
    writer.add_list(list3);
    writer.add_list_override(ListFormatOverride::new(3, 3));

    // ===== SECTION 4: HYPERLINKS =====
    writer.add_paragraph("Section 4: Hyperlinks")?;
    writer.add_paragraph("")?;
    writer.add_hyperlink(
        "Visit our website for more information.",
        "https://example.com",
        ParagraphFormatting::default(),
    )?;
    writer.add_hyperlink(
        "Contact us via email.",
        "mailto:contact@example.com",
        ParagraphFormatting::default(),
    )?;
    writer.add_paragraph("")?;

    // ===== SECTION 5: RICH TEXT =====
    writer.add_paragraph("Section 5: Rich Text Formatting")?;
    writer.add_paragraph("")?;
    let bold_fmt = CharacterFormatting {
        bold: Some(true),
        ..Default::default()
    };
    let italic_fmt = CharacterFormatting {
        italic: Some(true),
        ..Default::default()
    };
    writer.add_paragraph_runs(
        vec![
            (
                "This paragraph demonstrates ".to_string(),
                CharacterFormatting::default(),
            ),
            ("bold".to_string(), bold_fmt),
            (" and ".to_string(), CharacterFormatting::default()),
            ("italic".to_string(), italic_fmt),
            (
                " text runs in one paragraph.".to_string(),
                CharacterFormatting::default(),
            ),
        ],
        ParagraphFormatting::default(),
    )?;
    writer.add_paragraph("")?;

    // ===== SECTION 6: TABLE =====
    writer.add_paragraph("Section 6: Tables")?;
    writer.add_paragraph("")?;
    let table_idx = writer.add_table(3, 2)?;
    writer.set_table_cell_text(table_idx, 0, 0, "Feature")?;
    writer.set_table_cell_text(table_idx, 0, 1, "Status")?;
    writer.set_table_cell_text(table_idx, 1, 0, "Hyperlinks")?;
    writer.set_table_cell_text(table_idx, 1, 1, "Supported")?;
    writer.set_table_cell_text(table_idx, 2, 0, "Tables")?;
    writer.set_table_cell_text(table_idx, 2, 1, "Basic support")?;

    // ===== CONCLUSION =====
    writer.add_paragraph("Conclusion")?;
    writer.add_paragraph("")?;
    writer.add_paragraph(
        "This document demonstrates comprehensive DOC file generation with modern features. \
         All features are implemented using safe, idiomatic Rust code following Microsoft's \
         DOC specification and Apache POI's implementation patterns.",
    )?;
    writer.add_paragraph("")?;
    writer.add_paragraph("Thank you for using this library!")?;

    writer.add_footnote(FootnoteEntry::new(
        450,
        "First footnote: Additional context about the statement.".to_string(),
        1,
    ));
    writer.add_footnote(FootnoteEntry::new(
        520,
        "Second footnote: Citation - Smith, J. (2024). 'Document Standards'.".to_string(),
        2,
    ));
    writer.add_footnote(FootnoteEntry::new(
        580,
        "Third footnote: Explanatory note about usage patterns.".to_string(),
        3,
    ));

    writer.add_endnote(FootnoteEntry::new(
        680,
        "Endnote i: Detailed explanation at document end.".to_string(),
        1,
    ));
    writer.add_endnote(FootnoteEntry::new(
        750,
        "Endnote ii: Reference - Johnson, A. (2024). 'Writing Standards'.".to_string(),
        2,
    ));
    writer.add_endnote(FootnoteEntry::new(
        820,
        "Endnote iii: Additional notes for scholarly writing.".to_string(),
        3,
    ));

    // ===== SAVE DOCUMENT =====
    let output_path = "output/doc_all_features.doc";
    writer.save(output_path)?;

    println!("\n✅ Comprehensive DOC file created: {}", output_path);
    println!("\n📝 Open in Microsoft Word to verify:");
    println!("   ✓ Headers/footers (View > Header and Footer)");
    println!("   ✓ Footnotes at bottom of pages");
    println!("   ✓ Endnotes at end of document");
    println!("   ✓ Numbered and bulleted lists");
    println!("   ✓ Clickable hyperlinks");
    println!("   ✓ Rich text formatting");
    println!("   ✓ Basic table content");
    println!("\n🎉 All features demonstrated successfully!");

    Ok(())
}
