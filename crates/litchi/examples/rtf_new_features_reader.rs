//! Example demonstrating RTF reading with new features:
//! - Headers and Footers
//! - Footnotes and Endnotes
//! - Hyperlinks
//! - Track Changes (Revisions)

use litchi::rtf::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("RTF Document Feature Reader\n");
    println!("{}", "=".repeat(60));

    // Create a sample RTF document with all new features
    let rtf_content = r#"{\rtf1\ansi\deff0
{\fonttbl{\f0\fswiss Arial;}}
{\colortbl;\red0\green0\blue255;}

{\header{\pard\qc Header Text\par}}
{\footer{\pard\qc Footer Text - Page 1\par}}

{\fs32\b RTF Document with Advanced Features\par}
\par

{\fs24 This is a paragraph with a hyperlink: {\field{\*\fldinst{HYPERLINK "https://github.com"}}{\fldrslt{\ul\cf1 GitHub}}} for more info.\par}
\par

{\fs24 This paragraph has a footnote{\footnote{\chftn1 This is footnote content.}}.\par}
\par

{\fs24 This paragraph has an endnote{\endnote{\chftn1 This is endnote content.}}.\par}
\par

{\fs24 This text contains {\revised\revauth1 inserted content} with track changes.\par}
\par

}"#;

    println!("\n📄 Parsing RTF document...\n");

    // Parse the document
    let doc = RtfDocument::parse(rtf_content)?;

    // Display basic document info
    println!("✅ Document parsed successfully!\n");
    println!("📊 Document Statistics:");
    println!("  - Text length: {} characters", doc.text().len());
    println!("  - Paragraph count: {}", doc.paragraph_count());
    println!("  - Font count: {}", doc.font_table().fonts().len());
    println!("  - Color count: {}", doc.color_table().colors().len());

    // Display text content
    println!("\n📝 Document Text:");
    println!("{}", "-".repeat(60));
    println!("{}", doc.text());
    println!("{}", "-".repeat(60));

    // Display sections with headers/footers
    println!("\n📑 Sections and Headers/Footers:");
    println!("{}", "=".repeat(60));
    let sections = doc.sections();
    if sections.is_empty() {
        println!("  ℹ️  No sections found (using default section)");
    } else {
        for (i, section) in sections.iter().enumerate() {
            println!("\n  Section {}:", i + 1);
            println!(
                "    Page size: {}x{} twips",
                section.properties.page_width, section.properties.page_height
            );
            println!(
                "    Margins: L:{} R:{} T:{} B:{} twips",
                section.properties.margin_left,
                section.properties.margin_right,
                section.properties.margin_top,
                section.properties.margin_bottom
            );

            // Display headers and footers
            if section.headers_footers.is_empty() {
                println!("    No headers/footers");
            } else {
                for hf in &section.headers_footers {
                    let hf_type = match hf.header_type {
                        HeaderFooterType::Header => "Header",
                        HeaderFooterType::HeaderFirst => "First Page Header",
                        HeaderFooterType::HeaderLeft => "Left Page Header",
                        HeaderFooterType::HeaderRight => "Right Page Header",
                        HeaderFooterType::Footer => "Footer",
                        HeaderFooterType::FooterFirst => "First Page Footer",
                        HeaderFooterType::FooterLeft => "Left Page Footer",
                        HeaderFooterType::FooterRight => "Right Page Footer",
                    };
                    println!("    {} text: {}", hf_type, hf.text());
                }
            }
        }
    }

    // Display footnotes
    println!("\n📌 Footnotes:");
    println!("{}", "=".repeat(60));
    let footnotes = doc.footnotes();
    if footnotes.is_empty() {
        println!("  ℹ️  No footnotes found");
    } else {
        for (i, note) in footnotes.iter().enumerate() {
            println!("  [{}] Reference: {}", i + 1, note.reference);
            println!("      Content: {}", note.content);
            if note.formatting.italic {
                println!("      Style: Italic");
            }
        }
    }

    // Display endnotes
    println!("\n📋 Endnotes:");
    println!("{}", "=".repeat(60));
    let endnotes = doc.endnotes();
    if endnotes.is_empty() {
        println!("  ℹ️  No endnotes found");
    } else {
        for (i, note) in endnotes.iter().enumerate() {
            println!("  [{}] Reference: {}", i + 1, note.reference);
            println!("      Content: {}", note.content);
        }
    }

    // Display all notes together
    println!("\n📝 All Notes (Footnotes + Endnotes):");
    println!("{}", "=".repeat(60));
    let all_notes = doc.notes();
    if all_notes.is_empty() {
        println!("  ℹ️  No notes found");
    } else {
        for (i, note) in all_notes.iter().enumerate() {
            let note_type = if note.is_footnote {
                "Footnote"
            } else {
                "Endnote"
            };
            println!(
                "  {} {}: [{}] {}",
                i + 1,
                note_type,
                note.reference,
                note.content
            );
        }
    }

    // Display fields (including hyperlinks)
    println!("\n🔗 Fields and Hyperlinks:");
    println!("{}", "=".repeat(60));
    let fields = doc.fields();
    if fields.is_empty() {
        println!("  ℹ️  No fields found");
    } else {
        for (i, field) in fields.iter().enumerate() {
            let field_type_name = match field.field_type {
                FieldType::Hyperlink => "Hyperlink",
                FieldType::Reference => "Reference",
                FieldType::PageReference => "Page Reference",
                FieldType::NoteReference => "Note Reference",
                FieldType::Page => "Page",
                FieldType::Date => "Date",
                FieldType::Toc => "Table of Contents",
                FieldType::Bookmark => "Bookmark",
                FieldType::Equation => "Equation",
                FieldType::Index => "Index",
                FieldType::Unknown => "Unknown",
            };

            println!("  Field {}: {}", i + 1, field_type_name);
            println!("    Instruction: {}", field.instruction);
            if !field.result.is_empty() {
                println!("    Result: {}", field.result);
            }

            // Extract URL for hyperlinks
            if field.field_type == FieldType::Hyperlink
                && let Some(url) = field.extract_url()
            {
                println!("    URL: {}", url);
            }
        }
    }

    // Display track changes/revisions
    println!("\n✏️  Track Changes (Revisions):");
    println!("{}", "=".repeat(60));
    let revisions = doc.revisions();
    if revisions.is_empty() {
        println!("  ℹ️  No revisions found");
    } else {
        for (i, revision) in revisions.iter().enumerate() {
            let rev_type = match revision.revision_type {
                RevisionType::Insertion => "Insertion",
                RevisionType::Deletion => "Deletion",
                RevisionType::FormatChange => "Format Change",
                RevisionType::MovedFrom => "Moved From",
                RevisionType::MovedTo => "Moved To",
            };

            println!("  Revision {}: {}", i + 1, rev_type);
            println!("    Author: {}", revision.author);
            if let Some(ref date) = revision.date {
                println!("    Date: {}", date);
            }
            println!("    Content: {}", revision.content);
        }
    }

    // Display bookmarks
    println!("\n🔖 Bookmarks:");
    println!("{}", "=".repeat(60));
    let bookmarks = doc.bookmarks();
    if bookmarks.bookmarks().is_empty() {
        println!("  ℹ️  No bookmarks found");
    } else {
        for (i, bookmark) in bookmarks.bookmarks().iter().enumerate() {
            println!("  Bookmark {}: {}", i + 1, bookmark.name);
            if !bookmark.content.is_empty() {
                println!("    Content: {}", bookmark.content);
            }
        }
    }

    println!("\n{}", "=".repeat(60));
    println!("✅ Document analysis complete!");

    Ok(())
}
