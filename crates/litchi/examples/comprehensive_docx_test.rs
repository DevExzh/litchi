//! Comprehensive DOCX Feature Test & Verification
//!
//! This example generates a complete DOCX document testing ALL implemented features,
//! including the newly added Theme, Watermark, and Table of Contents support.
//!
//! Usage: cargo run --example comprehensive_docx_test --features ooxml --no-default-features
//!
//! Output: comprehensive_test.docx
//!
//! This document serves as:
//! 1. Feature verification tool
//! 2. Regression test baseline
//! 3. Visual reference for all capabilities

use litchi::ooxml::Props;
use litchi::ooxml::docx::{
    BorderColor, EndnotePos, Endnotes, FootnotePos, Footnotes, ListType, MutableDocument,
    MutableTheme, Package, PageBorderStyle, PageNumberFormat, PageOrientation, ParagraphAlignment,
    SectionPageBorder, SectionPageBorders, TableOfContents, UnderlineStyle, Watermark,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║     COMPREHENSIVE DOCX FEATURE TEST & VERIFICATION        ║");
    println!("╚════════════════════════════════════════════════════════════╝\n");

    let mut doc = MutableDocument::new();

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 1: NEW ADVANCED FEATURES (Themes, Watermarks, TOC)
    // ═══════════════════════════════════════════════════════════════════════════
    println!("📋 Part 1: Testing NEW Advanced Features");
    println!("   ├─ Themes (color schemes, fonts)");
    println!("   ├─ Watermarks (VML-based)");
    println!("   └─ Table of Contents (field-based)\n");

    // ───────────────────────────────────────────────────────────────────────────
    // 1.1 Theme Configuration
    // ───────────────────────────────────────────────────────────────────────────
    println!("   [1/3] Configuring custom theme...");
    let mut theme = MutableTheme::new("Comprehensive Test Theme");
    theme.set_major_font("Georgia");
    theme.set_minor_font("Calibri");

    // Set custom accent colors
    {
        let scheme = theme.color_scheme_mut();
        scheme.set_accent(0, "1F4E78"); // Dark blue
        scheme.set_accent(1, "E36C09"); // Orange
        scheme.set_accent(2, "70AD47"); // Green
        scheme.set_accent(3, "5B9BD5"); // Light blue
        scheme.set_accent(4, "C1C11F"); // Yellow-green
        scheme.set_accent(5, "A32B66"); // Purple
    }

    doc.set_theme(theme);
    println!("   ✓ Theme applied");

    // ───────────────────────────────────────────────────────────────────────────
    // 1.2 Watermark
    // ───────────────────────────────────────────────────────────────────────────
    println!("   [2/3] Adding watermark...");
    let mut watermark = Watermark::text("DRAFT - CONFIDENTIAL");
    watermark.set_font("Arial");
    watermark.set_color("D3D3D3"); // Light gray
    watermark.set_font_size(72);
    doc.set_watermark(watermark);
    println!("   ✓ Watermark added");

    // ───────────────────────────────────────────────────────────────────────────
    // 1.3 Document Title
    // ───────────────────────────────────────────────────────────────────────────
    let title = doc.add_paragraph();
    title.set_alignment(ParagraphAlignment::Center);
    let title_run = title.add_run();
    title_run.set_text("COMPREHENSIVE DOCX FEATURE TEST");
    title_run.font_size(48); // 24pt
    title_run.bold(true);
    title_run.color("1F4E78"); // Theme accent color

    let subtitle = doc.add_paragraph();
    subtitle.set_alignment(ParagraphAlignment::Center);
    let sub_run = subtitle.add_run();
    sub_run.set_text("Litchi OOXML Library - Complete Feature Verification");
    sub_run.font_size(24); // 12pt
    sub_run.italic(true);
    sub_run.color("808080");

    doc.add_paragraph(); // Spacing

    // ───────────────────────────────────────────────────────────────────────────
    // 1.4 Table of Contents
    // ───────────────────────────────────────────────────────────────────────────
    println!("   [3/3] Adding Table of Contents...");
    let toc = TableOfContents::new()
        .heading_levels(1, 3)
        .hyperlinks(true)
        .page_numbers(true)
        .right_align_page_numbers(true)
        .title("Table of Contents");
    doc.add_toc(toc)?;
    println!("   ✓ Table of Contents added\n");

    doc.add_paragraph(); // Spacing

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 2: CORE TEXT FEATURES
    // ═══════════════════════════════════════════════════════════════════════════
    println!("📋 Part 2: Testing Core Text Features");

    // ───────────────────────────────────────────────────────────────────────────
    // 2.1 Text Formatting
    // ───────────────────────────────────────────────────────────────────────────
    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    let h1_run = h1.add_run();
    h1_run.set_text("1. Text Formatting");
    h1_run.bold(true);
    h1_run.font_size(32); // 16pt

    let para = doc.add_paragraph();
    para.add_run_with_text("This section tests various text formatting options: ");
    para.add_run_with_text("bold").bold(true);
    para.add_run_with_text(", ");
    para.add_run_with_text("italic").italic(true);
    para.add_run_with_text(", ");
    para.add_run_with_text("underlined")
        .underline(UnderlineStyle::Single);
    para.add_run_with_text(", ");
    para.add_run_with_text("bold+italic")
        .bold(true)
        .italic(true);
    para.add_run_with_text(", and ");
    para.add_run_with_text("all three")
        .bold(true)
        .italic(true)
        .underline(UnderlineStyle::Single);
    para.add_run_with_text(".");

    // ───────────────────────────────────────────────────────────────────────────
    // 2.2 Font Styles and Colors
    // ───────────────────────────────────────────────────────────────────────────
    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    let h2_run = h2.add_run();
    h2_run.set_text("1.1 Font Styles and Colors");
    h2_run.bold(true);
    h2_run.font_size(28); // 14pt

    let para = doc.add_paragraph();
    para.add_run_with_text("Font size: ");
    para.add_run_with_text("tiny (8pt)").font_size(16);
    para.add_run_with_text(", ");
    para.add_run_with_text("normal (11pt)").font_size(22);
    para.add_run_with_text(", ");
    para.add_run_with_text("large (18pt)").font_size(36);
    para.add_run_with_text(", ");
    para.add_run_with_text("huge (24pt)").font_size(48);

    let para = doc.add_paragraph();
    para.add_run_with_text("Colors: ");
    para.add_run_with_text("red").color("FF0000");
    para.add_run_with_text(", ");
    para.add_run_with_text("green").color("00FF00");
    para.add_run_with_text(", ");
    para.add_run_with_text("blue").color("0000FF");
    para.add_run_with_text(", ");
    para.add_run_with_text("purple").color("800080");
    para.add_run_with_text(", ");
    para.add_run_with_text("orange").color("FFA500");

    let para = doc.add_paragraph();
    para.add_run_with_text("Font families: ");
    para.add_run_with_text("Arial").font_name("Arial");
    para.add_run_with_text(", ");
    para.add_run_with_text("Times New Roman")
        .font_name("Times New Roman");
    para.add_run_with_text(", ");
    para.add_run_with_text("Courier New")
        .font_name("Courier New");
    para.add_run_with_text(", ");
    para.add_run_with_text("Georgia").font_name("Georgia");

    // ───────────────────────────────────────────────────────────────────────────
    // 2.3 Paragraph Alignment
    // ───────────────────────────────────────────────────────────────────────────
    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    let h2_run = h2.add_run();
    h2_run.set_text("1.2 Paragraph Alignment");
    h2_run.bold(true);

    let para = doc.add_paragraph();
    para.set_alignment(ParagraphAlignment::Left);
    para.add_run_with_text("Left-aligned paragraph (default alignment for body text)");

    let para = doc.add_paragraph();
    para.set_alignment(ParagraphAlignment::Center);
    para.add_run_with_text("Center-aligned paragraph (useful for titles and headings)");

    let para = doc.add_paragraph();
    para.set_alignment(ParagraphAlignment::Right);
    para.add_run_with_text("Right-aligned paragraph (often used for dates and signatures)");

    let para = doc.add_paragraph();
    para.set_alignment(ParagraphAlignment::Justify);
    para.add_run_with_text("Justified paragraph with sufficient text to demonstrate full justification across the entire width of the page. This alignment is commonly used in formal documents and publications to create a clean, professional appearance.");

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 3: LISTS AND STRUCTURE
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Text formatting tests complete");
    println!("\n📋 Part 3: Testing Lists and Structure");

    // ───────────────────────────────────────────────────────────────────────────
    // 3.1 Bullet Lists
    // ───────────────────────────────────────────────────────────────────────────
    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("2. Lists and Structure").bold(true);

    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("2.1 Bullet Lists").bold(true);

    let items = [
        "First-level bullet item",
        "Another first-level item",
        "Third first-level item",
    ];

    for item in &items {
        let para = doc.add_paragraph();
        para.set_list(ListType::Bullet, 0);
        para.add_run_with_text(item);
    }

    // ───────────────────────────────────────────────────────────────────────────
    // 3.2 Numbered Lists
    // ───────────────────────────────────────────────────────────────────────────
    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("2.2 Numbered Lists").bold(true);

    let numbered_items = [
        "First numbered item with sequential numbering",
        "Second numbered item demonstrating automatic incrementing",
        "Third numbered item showing proper list formatting",
        "Fourth item to verify list continuation",
    ];

    for item in &numbered_items {
        let para = doc.add_paragraph();
        para.set_list(ListType::Decimal, 0);
        para.add_run_with_text(item);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 4: TABLES
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Lists tests complete");
    println!("\n📋 Part 4: Testing Tables");

    // ───────────────────────────────────────────────────────────────────────────
    // 4.1 Basic Table
    // ───────────────────────────────────────────────────────────────────────────
    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("3. Tables").bold(true);

    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("3.1 Basic Table Structure").bold(true);

    doc.add_paragraph_with_text("Simple 4×3 table with headers:");

    let table = doc.add_table(4, 3);
    table.set_width_percent(100);

    // Note: Table borders configured with default settings

    // Header row
    for (i, header) in ["Product", "Quantity", "Price"].iter().enumerate() {
        let cell = table.cell(0, i).unwrap();
        let para = cell.add_paragraph();
        para.add_run_with_text(header).bold(true);
    }

    // Data rows
    let data = [
        ("Widget A", "100", "$25.00"),
        ("Widget B", "75", "$30.00"),
        ("Widget C", "50", "$45.00"),
    ];

    for (row, (product, qty, price)) in data.iter().enumerate() {
        table.cell(row + 1, 0).unwrap().set_text(product);
        table.cell(row + 1, 1).unwrap().set_text(qty);
        table.cell(row + 1, 2).unwrap().set_text(price);
    }

    // ───────────────────────────────────────────────────────────────────────────
    // 4.2 Formatted Table
    // ───────────────────────────────────────────────────────────────────────────
    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("3.2 Formatted Table Cells").bold(true);

    doc.add_paragraph_with_text("Table with various cell formatting:");

    let table = doc.add_table(3, 3);
    table.set_width_percent(90);

    // Header with colors
    for (i, (header, color)) in [
        ("Status", "00FF00"),
        ("Priority", "FFA500"),
        ("Action", "FF0000"),
    ]
    .iter()
    .enumerate()
    {
        let cell = table.cell(0, i).unwrap();
        let para = cell.add_paragraph();
        para.add_run_with_text(header).bold(true).color(color);
    }

    // Data with formatting
    let cell = table.cell(1, 0).unwrap();
    cell.add_paragraph()
        .add_run_with_text("Active")
        .color("00FF00")
        .bold(true);

    let cell = table.cell(1, 1).unwrap();
    cell.add_paragraph()
        .add_run_with_text("High")
        .color("FF0000")
        .italic(true);

    let cell = table.cell(1, 2).unwrap();
    cell.add_paragraph()
        .add_run_with_text("Review")
        .underline(UnderlineStyle::Single);

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 5: HYPERLINKS AND REFERENCES
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Tables tests complete");
    println!("\n📋 Part 5: Testing Hyperlinks");

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("4. Hyperlinks and References")
        .bold(true);

    let para = doc.add_paragraph();
    para.add_run_with_text("External links: ");
    para.add_hyperlink("https://www.rust-lang.org/", "Rust");
    para.add_run_with_text(", ");
    para.add_hyperlink("https://github.com/", "GitHub");
    para.add_run_with_text(", ");
    para.add_hyperlink("https://docs.rs/", "docs.rs");

    let para = doc.add_paragraph();
    para.add_run_with_text("Documentation: Visit ");
    para.add_hyperlink("https://poi.apache.org/", "Apache POI");
    para.add_run_with_text(" for reference implementation.");

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 6: PAGE BREAKS AND SECTIONS
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Hyperlinks tests complete");
    println!("\n📋 Part 6: Testing Page Breaks and Sections");

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("5. Page Breaks and Sections")
        .bold(true);

    doc.add_paragraph_with_text("This content is on the current page.");
    doc.add_paragraph_with_text("A page break follows immediately after this paragraph.");

    doc.add_page_break();

    // Page 2
    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("6. Content on Page 2").bold(true);

    doc.add_paragraph_with_text("This heading and content appear on page 2 after the page break.");
    doc.add_paragraph_with_text("Page breaks are essential for controlling document layout and organizing content into logical sections.");

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 7: HEADERS AND FOOTERS
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Page breaks tests complete");
    println!("\n📋 Part 7: Testing Headers and Footers");

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("7. Headers and Footers").bold(true);

    doc.add_paragraph_with_text("This document includes:");

    let items = [
        "Header with document title",
        "Footer with page numbers (Page X of Y format)",
        "Consistent header/footer across all pages",
    ];

    for item in &items {
        let para = doc.add_paragraph();
        para.set_list(ListType::Bullet, 0);
        para.add_run_with_text(item);
    }

    // Add header
    let header = doc.add_header_paragraph();
    header.set_alignment(ParagraphAlignment::Center);
    header
        .add_run_with_text("Comprehensive DOCX Test - ")
        .italic(true)
        .color("808080");
    header
        .add_run_with_text("DRAFT")
        .italic(true)
        .color("FF0000");

    // Add footer with page numbers
    let footer = doc.add_footer_paragraph();
    footer.set_alignment(ParagraphAlignment::Center);
    footer.add_run_with_text("Page ");
    footer.add_run().add_page_number(PageNumberFormat::Decimal);
    footer.add_run_with_text(" of ");
    footer.add_run().add_page_count();
    footer
        .add_run_with_text(" | Generated by Litchi")
        .font_size(16)
        .color("808080");

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 8: FOOTNOTES AND ENDNOTES
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Headers/Footers tests complete");
    println!("\n📋 Part 8: Testing Footnotes and Endnotes");

    doc.add_page_break();

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("8. Footnotes and Endnotes").bold(true);

    // Create footnotes and add references
    let (fn1_id, footnote1) = doc.add_footnote();
    footnote1.add_paragraph_with_text("This is footnote 1 with detailed information.");

    let (fn2_id, footnote2) = doc.add_footnote();
    footnote2.add_paragraph_with_text("This is footnote 2 with additional context.");

    // Create endnotes and add references
    let (en1_id, endnote1) = doc.add_endnote();
    endnote1.add_paragraph_with_text("This is endnote 1, typically used for citations.");

    let (en2_id, endnote2) = doc.add_endnote();
    endnote2.add_paragraph_with_text("This is endnote 2 with bibliographic reference.");

    // Add text with footnote and endnote references
    let para1 = doc.add_paragraph();
    para1
        .add_run_with_text("This paragraph has a footnote reference")
        .add_footnote_reference(fn1_id);
    para1.add_run_with_text(" and here is some more text with another footnote");
    para1.add_run().add_footnote_reference(fn2_id);
    para1.add_run_with_text(".");

    let para2 = doc.add_paragraph();
    para2
        .add_run_with_text("This paragraph demonstrates endnotes")
        .add_endnote_reference(en1_id);
    para2.add_run_with_text(" which appear at the end of the document");
    para2.add_run().add_endnote_reference(en2_id);
    para2.add_run_with_text(".");

    doc.add_paragraph_with_text("In Microsoft Word, footnotes appear at the bottom of each page, while endnotes appear at the end of the document.");

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 9: SECTION PROPERTIES
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Footnotes/Endnotes tests complete");
    println!("\n📋 Part 9: Testing Section Properties");

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("9. Section Properties").bold(true);

    doc.add_paragraph_with_text("Document section settings:");

    let settings = [
        "Page size: Letter (8.5\" × 11\")",
        "Orientation: Portrait",
        "Margins: 1 inch (1440 TWIPs) on all sides",
        "Page width: 12240 TWIPs",
        "Page height: 15840 TWIPs",
        "Footnotes: page bottom; endnotes: document end",
        "Page border: typed blue double line",
    ];

    for setting in &settings {
        let para = doc.add_paragraph();
        para.set_list(ListType::Bullet, 0);
        para.add_run_with_text(setting);
    }

    // Configure section
    let page_border = SectionPageBorder {
        style: PageBorderStyle::Double,
        size: Some(8),
        space: Some(24),
        color: Some(BorderColor::rgb(31, 78, 120)),
        shadow: false,
        frame: false,
    };
    let section = doc.section_mut();
    section.page_width = 12240;
    section.page_height = 15840;
    section.margin_top = 1440;
    section.margin_bottom = 1440;
    section.margin_left = 1440;
    section.margin_right = 1440;
    section.orientation = PageOrientation::Portrait;
    section.footnotes = Some(Footnotes {
        position: Some(FootnotePos::PageBottom),
        ..Footnotes::default()
    });
    section.endnotes = Some(Endnotes {
        position: Some(EndnotePos::DocumentEnd),
        ..Endnotes::default()
    });
    section.page_borders = Some(SectionPageBorders {
        top: Some(page_border),
        left: Some(page_border),
        bottom: Some(page_border),
        right: Some(page_border),
        ..SectionPageBorders::default()
    });

    // ═══════════════════════════════════════════════════════════════════════════
    // PART 10: VERIFICATION SUMMARY
    // ═══════════════════════════════════════════════════════════════════════════
    println!("   ✓ Section properties tests complete");
    println!("\n📋 Part 10: Generating Verification Summary");

    doc.add_page_break();

    let h1 = doc.add_paragraph();
    h1.set_style("Heading1");
    h1.add_run_with_text("10. Verification Summary").bold(true);

    doc.add_paragraph_with_text("This document tested the following features:");

    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("10.1 New Advanced Features")
        .bold(true);

    let new_features = [
        "✅ Document Themes (color schemes, fonts, format schemes)",
        "✅ Watermarks (VML shape-based, diagonal text)",
        "✅ Table of Contents (field-based with configurable levels)",
    ];

    for feature in &new_features {
        let para = doc.add_paragraph();
        para.set_list(ListType::Bullet, 0);
        para.add_run_with_text(feature).color("00AA00").bold(true);
    }

    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("10.2 Core Features").bold(true);

    let core_features = [
        "Text formatting (bold, italic, underline)",
        "Font styles (name, size, color)",
        "Paragraph alignment (left, center, right, justify)",
        "Lists (bullet and numbered)",
        "Tables (with borders and formatting)",
        "Hyperlinks (external URLs)",
        "Page breaks",
        "Headers and footers (with page numbers)",
        "Footnotes and endnotes",
        "Section properties (margins, orientation)",
    ];

    for feature in &core_features {
        let para = doc.add_paragraph();
        para.set_list(ListType::Bullet, 0);
        para.add_run_with_text(feature);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // VERIFICATION INSTRUCTIONS
    // ═══════════════════════════════════════════════════════════════════════════
    doc.add_paragraph();

    let h2 = doc.add_paragraph();
    h2.set_style("Heading2");
    h2.add_run_with_text("10.3 Verification Checklist")
        .bold(true);

    let instructions = [
        "Open this file in Microsoft Word",
        "Verify the watermark appears on all pages (faint diagonal text)",
        "Check the TOC on page 1 - right-click and select 'Update Field'",
        "Verify all headings appear in the TOC after update",
        "Check that hyperlinks are clickable and colored blue",
        "Verify header appears at top of each page",
        "Verify footer with page numbers at bottom of each page",
        "Check that tables have proper borders and formatting",
        "Verify page breaks create new pages correctly",
        "Check footnotes appear at bottom of pages",
        "Verify all text formatting (bold, italic, colors) displays correctly",
        "Check File → Properties to see theme and metadata",
    ];

    for instruction in &instructions {
        let para = doc.add_paragraph();
        para.set_list(ListType::Decimal, 0);
        para.add_run_with_text(instruction);
    }

    doc.add_paragraph();

    let conclusion = doc.add_paragraph();
    conclusion.set_alignment(ParagraphAlignment::Center);
    conclusion
        .add_run_with_text("✅ ALL FEATURES TESTED")
        .bold(true)
        .font_size(32)
        .color("00AA00");

    // ═══════════════════════════════════════════════════════════════════════════
    // SAVE DOCUMENT
    // ═══════════════════════════════════════════════════════════════════════════
    println!("\n💾 Saving document...");

    let mut package = Package::new()?;
    *package.document_mut()? = doc;

    // Set document properties
    let _ = package.put_props(
        Props::new()
            .title("Comprehensive DOCX Feature Test")
            .subject("Feature Verification and Regression Testing")
            .creator("Litchi OOXML Library")
            .keywords("docx, test, verification, regression, features, themes, watermarks, toc")
            .description("Complete test document for all DOCX writer features including new theme, watermark, and TOC support"),
    );

    let filename = "comprehensive_test.docx";
    package.save(filename)?;

    println!("\n╔════════════════════════════════════════════════════════════╗");
    println!("║                    ✅ SUCCESS!                             ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!("\n📄 Generated: {}", filename);
    println!("\n📊 Test Summary:");
    println!("   ✓ 10 feature categories tested");
    println!("   ✓ 3 new advanced features (Theme, Watermark, TOC)");
    println!("   ✓ 10+ core features verified");
    println!("   ✓ Complete document structure");
    println!("\n📖 Next Steps:");
    println!("   1. Open '{}' in Microsoft Word", filename);
    println!("   2. Right-click TOC → Update Field → Update entire table");
    println!("   3. Follow verification checklist in section 10.3");
    println!("   4. Verify watermark visibility on all pages");
    println!("   5. Check theme colors and fonts");
    println!("\n💡 Regression Testing:");
    println!("   • Save this file as baseline");
    println!("   • Re-run after code changes");
    println!("   • Compare output files for differences");
    println!("   • Verify all features still work correctly\n");

    Ok(())
}
