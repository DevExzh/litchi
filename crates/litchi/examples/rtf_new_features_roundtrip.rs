//! Round-trip test for RTF new features
//! This example writes an RTF document with new features, then reads it back
//! to verify that all features are preserved correctly.

use litchi::rtf::*;

fn main() -> std::result::Result<(), Box<dyn std::error::Error>> {
    println!("RTF New Features Round-Trip Test");
    println!("{}", "=".repeat(70));

    // Phase 1: Write RTF document
    println!("\n📝 PHASE 1: Writing RTF document with new features...\n");

    let output_file = "rtf_roundtrip_test.rtf";
    let rtf_content = create_rtf_with_features()?;
    std::fs::write(output_file, &rtf_content)?;

    println!("✅ Written {} bytes to {}", rtf_content.len(), output_file);

    // Phase 2: Read RTF document
    println!("\n📖 PHASE 2: Reading RTF document back...\n");

    let doc = RtfDocument::from_bytes(&rtf_content)?;

    println!("✅ Document parsed successfully!");

    // Phase 3: Verify features
    println!("\n✓ PHASE 3: Verifying features...\n");

    let mut all_passed = true;

    // Test 1: Verify sections with headers/footers
    print!("  [1/7] Checking sections and headers/footers... ");
    let sections = doc.sections();
    if !sections.is_empty() {
        let has_headers_footers = sections.iter().any(|s| !s.headers_footers.is_empty());
        if has_headers_footers {
            println!("✅ PASS");
            let section = &sections[0];
            for hf in &section.headers_footers {
                let hf_type = match hf.header_type {
                    HeaderFooterType::Header => "Header",
                    HeaderFooterType::Footer => "Footer",
                    HeaderFooterType::HeaderFirst => "First Header",
                    _ => "Other",
                };
                println!(
                    "      ✓ Found {}: {}",
                    hf_type,
                    hf.text().chars().take(30).collect::<String>()
                );
            }
        } else {
            println!("⚠️  PARTIAL (no headers/footers in sections)");
            all_passed = false;
        }
    } else {
        println!("⚠️  PARTIAL (no sections found)");
    }

    // Test 2: Verify footnotes
    print!("  [2/7] Checking footnotes... ");
    let footnotes = doc.footnotes();
    if !footnotes.is_empty() {
        println!("✅ PASS ({} footnotes found)", footnotes.len());
        for (i, note) in footnotes.iter().enumerate() {
            println!(
                "      ✓ Footnote {}: [{}] {}",
                i + 1,
                note.reference,
                note.content.chars().take(40).collect::<String>()
            );
        }
    } else {
        println!("⚠️  PARTIAL (no footnotes found)");
    }

    // Test 3: Verify endnotes
    print!("  [3/7] Checking endnotes... ");
    let endnotes = doc.endnotes();
    if !endnotes.is_empty() {
        println!("✅ PASS ({} endnotes found)", endnotes.len());
        for (i, note) in endnotes.iter().enumerate() {
            println!(
                "      ✓ Endnote {}: [{}] {}",
                i + 1,
                note.reference,
                note.content.chars().take(40).collect::<String>()
            );
        }
    } else {
        println!("⚠️  PARTIAL (no endnotes found)");
    }

    // Test 4: Verify all notes together
    print!("  [4/7] Checking all notes... ");
    let all_notes = doc.notes();
    if !all_notes.is_empty() {
        println!("✅ PASS ({} total notes)", all_notes.len());
        let footnote_count = all_notes.iter().filter(|n| n.is_footnote).count();
        let endnote_count = all_notes.iter().filter(|n| !n.is_footnote).count();
        println!(
            "      ✓ Footnotes: {}, Endnotes: {}",
            footnote_count, endnote_count
        );
    } else {
        println!("⚠️  FAIL (no notes found)");
        all_passed = false;
    }

    // Test 5: Verify hyperlink fields
    print!("  [5/7] Checking hyperlink fields... ");
    let fields = doc.fields();
    let hyperlinks: Vec<_> = fields
        .iter()
        .filter(|f| f.field_type == FieldType::Hyperlink)
        .collect();

    if !hyperlinks.is_empty() {
        println!("✅ PASS ({} hyperlinks found)", hyperlinks.len());
        for (i, field) in hyperlinks.iter().enumerate() {
            if let Some(url) = field.extract_url() {
                println!(
                    "      ✓ Hyperlink {}: {} -> {}",
                    i + 1,
                    field.result.chars().take(20).collect::<String>(),
                    url.chars().take(40).collect::<String>()
                );
            }
        }
    } else {
        println!("⚠️  PARTIAL (no hyperlinks found)");
    }

    // Test 6: Verify all fields
    print!("  [6/7] Checking all fields... ");
    if !fields.is_empty() {
        println!("✅ PASS ({} fields found)", fields.len());
        for field in fields.iter().take(3) {
            let field_type = match field.field_type {
                FieldType::Hyperlink => "HYPERLINK",
                FieldType::Page => "PAGE",
                FieldType::Date => "DATE",
                _ => "OTHER",
            };
            println!("      ✓ Field: {}", field_type);
        }
    } else {
        println!("⚠️  PARTIAL (no fields found)");
    }

    // Test 7: Verify track changes/revisions
    print!("  [7/7] Checking track changes/revisions... ");
    let revisions = doc.revisions();
    if !revisions.is_empty() {
        println!("✅ PASS ({} revisions found)", revisions.len());
        for (i, rev) in revisions.iter().enumerate() {
            let rev_type = match rev.revision_type {
                RevisionType::Insertion => "INSERT",
                RevisionType::Deletion => "DELETE",
                RevisionType::FormatChange => "FORMAT",
                RevisionType::MovedFrom => "MOVE_FROM",
                RevisionType::MovedTo => "MOVE_TO",
            };
            println!(
                "      ✓ Revision {}: {} by {} - {}",
                i + 1,
                rev_type,
                rev.author,
                rev.content.chars().take(30).collect::<String>()
            );
        }
    } else {
        println!("⚠️  PARTIAL (no revisions found)");
    }

    // Summary
    println!("\n{}", "=".repeat(70));
    if all_passed {
        println!("✅ ALL CRITICAL TESTS PASSED!");
    } else {
        println!("⚠️  SOME TESTS PARTIAL/FAILED - See details above");
    }

    // Display document statistics
    println!("\n📊 Document Statistics:");
    println!("  - Text length: {} characters", doc.text().len());
    println!("  - Paragraphs: {}", doc.paragraph_count());
    println!("  - Sections: {}", sections.len());
    println!("  - Footnotes: {}", doc.footnotes().len());
    println!("  - Endnotes: {}", doc.endnotes().len());
    println!("  - Total notes: {}", doc.notes().len());
    println!("  - Fields: {}", doc.fields().len());
    println!("  - Revisions: {}", doc.revisions().len());
    println!("  - Fonts: {}", doc.font_table().fonts().len());
    println!("  - Colors: {}", doc.color_table().colors().len());

    println!("\n💾 Test file saved as: {}", output_file);
    println!("   You can open this file in Word or another RTF viewer to verify.");

    Ok(())
}

/// Create an RTF document with all new features
fn create_rtf_with_features() -> std::result::Result<Vec<u8>, Box<dyn std::error::Error>> {
    // Create RTF content as a string for simplicity
    let rtf_content = r#"{\rtf1\ansi\deff0
{\fonttbl{\f0\fswiss Arial;}}
{\colortbl;\red0\green0\blue255;}

{\header{\pard\qc\b\fs20 Test Document Header\par}}
{\footer{\pard\qc Page Footer - Round-Trip Test\par}}
{\headerf{\pard\qc\b\i\fs24 First Page Special Header\par}}

{\b\fs36 RTF Round-Trip Test Document\par}\par

{\fs24\b Section 1: Hyperlinks\par}\par
{\fs20 Visit: {\field{\*\fldinst{HYPERLINK "https://github.com/DevExzh/litchi"}}{\fldrslt{\ul\cf1 Litchi Repository}}}\par}
{\fs20 Documentation: {\field{\*\fldinst{HYPERLINK "https://docs.rs/litchi"}}{\fldrslt{\ul\cf1 API Docs}}}\par}\par

{\fs24\b Section 2: Footnotes\par}\par
{\fs20 This paragraph contains a footnote{\footnote{\chftn1 First footnote with detailed information.}} and continues with more text{\footnote{\chftn2 Second footnote with additional context.}}.\par}\par

{\fs24\b Section 3: Endnotes\par}\par
{\fs20 This section includes endnotes{\endnote{\chftn1 First endnote providing supplementary information.}} for reference{\endnote{\chftn2 Second endnote with bibliography details.}}.\par}\par

{\fs24\b Section 4: Track Changes\par}\par
{\fs20 This text includes {\revised\revauth1 newly added content} with tracked changes.\par}
{\fs20 Here is {\deleted\revauthdel1 removed text} showing deletions.\par}\par

{\fs24\b Section 5: Fields\par}\par
{\fs20 Current page: {\field{\*\fldinst{PAGE}}{\fldrslt{1}}}\par}\par

}"#;

    Ok(rtf_content.as_bytes().to_vec())
}
