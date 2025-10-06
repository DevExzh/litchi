//! Example demonstrating Markdown conversion for Office documents.
//!
//! This example shows how to convert Word documents and PowerPoint presentations
//! to Markdown format using the `ToMarkdown` trait.
//!
//! Run with:
//! ```sh
//! cargo run --example markdown_conversion
//! ```

use litchi::markdown::{MarkdownOptions, TableStyle, ToMarkdown};
use litchi::{Document, Presentation};

fn main() -> Result<(), litchi::Error> {
    println!("=== Litchi Markdown Conversion Example ===\n");

    // Example 1: Convert a document with default options
    println!("1. Converting document with default options...");
    if let Ok(doc) = Document::open("test.docx") {
        let markdown = doc.to_markdown()?;
        println!("Document markdown (first 200 chars):");
        println!("{}\n", &markdown.chars().take(200).collect::<String>());
    } else {
        println!("  (test.docx not found, skipping)\n");
    }

    // Example 2: Convert a document with custom options
    println!("2. Converting document with custom options...");
    if let Ok(doc) = Document::open("test.doc") {
        let options = MarkdownOptions::new()
            .with_styles(true) // Include bold, italic, etc.
            .with_metadata(false) // Don't include metadata
            .with_table_style(TableStyle::MinimalHtml); // Use HTML for tables

        let markdown = doc.to_markdown_with_options(&options)?;
        println!("Document markdown with HTML tables (first 200 chars):");
        println!("{}\n", &markdown.chars().take(200).collect::<String>());
    } else {
        println!("  (test.doc not found, skipping)\n");
    }

    // Example 3: Convert a presentation
    println!("3. Converting presentation...");
    if let Ok(pres) = Presentation::open("test.pptx") {
        let markdown = pres.to_markdown()?;
        println!("Presentation markdown (first 300 chars):");
        println!("{}\n", &markdown.chars().take(300).collect::<String>());
    } else {
        println!("  (test.pptx not found, skipping)\n");
    }

    // Example 4: Convert individual paragraphs
    println!("4. Converting individual paragraphs...");
    if let Ok(doc) = Document::open("test.docx") {
        let paragraphs = doc.paragraphs()?;
        println!("First 3 paragraphs:");
        for (i, para) in paragraphs.iter().take(3).enumerate() {
            let md = para.to_markdown()?;
            if !md.trim().is_empty() {
                println!("  Para {}: {}", i + 1, md.trim());
            }
        }
        println!();
    } else {
        println!("  (test.docx not found, skipping)\n");
    }

    // Example 5: Convert tables with different styles
    println!("5. Converting tables with different styles...");
    if let Ok(doc) = Document::open("test.docx") {
        let tables = doc.tables()?;
        if !tables.is_empty() {
            let table = &tables[0];

            // Markdown table style
            let options_md = MarkdownOptions::new().with_table_style(TableStyle::Markdown);
            let md_table = table.to_markdown_with_options(&options_md)?;
            println!("Markdown table style:");
            println!("{}\n", md_table);

            // Minimal HTML table style
            let options_html = MarkdownOptions::new().with_table_style(TableStyle::MinimalHtml);
            let html_table = table.to_markdown_with_options(&options_html)?;
            println!("Minimal HTML table style:");
            println!("{}\n", html_table);
        } else {
            println!("  No tables found in document\n");
        }
    } else {
        println!("  (test.docx not found, skipping)\n");
    }

    // Example 6: Working with different table indent
    println!("6. Styled HTML table with custom indent...");
    if let Ok(doc) = Document::open("test.docx") {
        let tables = doc.tables()?;
        if !tables.is_empty() {
            let table = &tables[0];

            let options = MarkdownOptions::new()
                .with_table_style(TableStyle::StyledHtml)
                .with_html_table_indent(4); // 4 spaces indent

            let styled_table = table.to_markdown_with_options(&options)?;
            println!("{}\n", styled_table);
        }
    } else {
        println!("  (test.docx not found, skipping)\n");
    }

    println!("\n=== Example Complete ===");
    Ok(())
}

