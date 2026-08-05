//! Example demonstrating list numbering in DOC files
//!
//! This example creates a DOC file with:
//! - Numbered lists with different formats
//! - Bulleted lists
//! - Multi-level nested lists
//!
//! Run with: cargo run --example doc_writer_numbering

use litchi_doc::writer::{
    DocWriter, ListFormatOverride, ListLevel, ListStructure, NumberFormat, ParagraphFormatting,
};
use std::error::Error;

/// Helper to create a list-associated paragraph
fn list_para(_text: &str, ilfo: u16, ilvl: u8) -> ParagraphFormatting {
    ParagraphFormatting {
        ilfo: Some(ilfo),
        ilvl: Some(ilvl),
        left_indent: Some(720 * (ilvl as i32 + 1)), // 0.5in per level
        first_line_indent: Some(-360),              // hanging indent
        ..Default::default()
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating DOC file with numbered and bulleted lists...");

    let mut writer = DocWriter::new();

    // Title
    writer.add_paragraph("List Numbering Examples")?;
    writer.add_paragraph("")?;

    // Section 1: Decimal numbering (ilfo=1)
    writer.add_paragraph("1. Decimal Numbered List")?;
    writer.add_formatted_paragraph("First item in decimal list", list_para("", 1, 0))?;
    writer.add_formatted_paragraph("Second item in decimal list", list_para("", 1, 0))?;
    writer.add_formatted_paragraph("Third item in decimal list", list_para("", 1, 0))?;
    writer.add_paragraph("")?;

    // Section 2: Roman numerals (ilfo=2)
    writer.add_paragraph("2. Roman Numeral List")?;
    writer.add_formatted_paragraph("First item in Roman numeral list", list_para("", 2, 0))?;
    writer.add_formatted_paragraph("Second item in Roman numeral list", list_para("", 2, 0))?;
    writer.add_formatted_paragraph("Third item in Roman numeral list", list_para("", 2, 0))?;
    writer.add_paragraph("")?;

    // Section 3: Letters (ilfo=3)
    writer.add_paragraph("3. Alphabetic List")?;
    writer.add_formatted_paragraph("First item in alphabetic list", list_para("", 3, 0))?;
    writer.add_formatted_paragraph("Second item in alphabetic list", list_para("", 3, 0))?;
    writer.add_formatted_paragraph("Third item in alphabetic list", list_para("", 3, 0))?;
    writer.add_paragraph("")?;

    // Section 4: Bullets (ilfo=4)
    writer.add_paragraph("4. Bulleted List")?;
    writer.add_formatted_paragraph("First bullet point", list_para("", 4, 0))?;
    writer.add_formatted_paragraph("Second bullet point", list_para("", 4, 0))?;
    writer.add_formatted_paragraph("Third bullet point", list_para("", 4, 0))?;
    writer.add_paragraph("")?;

    // Section 5: Multi-level list (ilfo=5)
    writer.add_paragraph("5. Multi-Level List")?;
    writer.add_formatted_paragraph("Level 1 - Item 1", list_para("", 5, 0))?;
    writer.add_formatted_paragraph("Level 2 - Item 1.1", list_para("", 5, 1))?;
    writer.add_formatted_paragraph("Level 2 - Item 1.2", list_para("", 5, 1))?;
    writer.add_formatted_paragraph("Level 3 - Item 1.2.1", list_para("", 5, 2))?;
    writer.add_formatted_paragraph("Level 3 - Item 1.2.2", list_para("", 5, 2))?;
    writer.add_formatted_paragraph("Level 1 - Item 2", list_para("", 5, 0))?;
    writer.add_formatted_paragraph("Level 2 - Item 2.1", list_para("", 5, 1))?;

    // List definitions
    let mut list1 = ListStructure::new(1);
    list1.add_level(ListLevel::new(1, NumberFormat::Decimal));
    writer.add_list(list1);
    writer.add_list_override(ListFormatOverride::new(1, 1));

    let mut list2 = ListStructure::new(2);
    list2.add_level(ListLevel::new(1, NumberFormat::UpperRoman));
    writer.add_list(list2);
    writer.add_list_override(ListFormatOverride::new(2, 2));

    let mut list3 = ListStructure::new(3);
    list3.add_level(ListLevel::new(1, NumberFormat::UpperLetter));
    writer.add_list(list3);
    writer.add_list_override(ListFormatOverride::new(3, 3));

    let mut list4 = ListStructure::new(4);
    list4.add_level(ListLevel::new(1, NumberFormat::Bullet));
    writer.add_list(list4);
    writer.add_list_override(ListFormatOverride::new(4, 4));

    let mut list5 = ListStructure::new(5);
    list5.add_level(ListLevel::new(1, NumberFormat::Decimal));
    list5.add_level(ListLevel::new(1, NumberFormat::LowerLetter));
    list5.add_level(ListLevel::new(1, NumberFormat::LowerRoman));
    writer.add_list(list5);
    writer.add_list_override(ListFormatOverride::new(5, 5));

    println!("List structures configured:");
    println!("  ✓ List 1: Decimal numbering (1, 2, 3...)");
    println!("  ✓ List 2: Roman numerals (I, II, III...)");
    println!("  ✓ List 3: Alphabetic (A, B, C...)");
    println!("  ✓ List 4: Bullets (•)");
    println!("  ✓ List 5: Multi-level (1, a, i...)");

    // Save the document
    let output_path = "output/doc_numbering.doc";
    writer.save(output_path)?;

    println!("\n✅ Document saved to: {}", output_path);
    println!("\n📝 Open this file in Microsoft Word to verify:");
    println!("   - Five different list types should be visible");
    println!("   - Home > Paragraph > Numbering to see list formats");
    println!("   - Multi-level list should show nested indentation");
    println!("   - Bullets should display as bullet points");

    Ok(())
}
