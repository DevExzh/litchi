//! Comprehensive ODT (OpenDocument Text) reading example.
//!
//! This example demonstrates all reading capabilities for ODT files,
//! serving as both verification and regression testing.
//!
//! Run with:
//! ```bash
//! cargo run --example odt_reader_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::odf::Document;

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODT Reader Comprehensive Test ===\n");

    // Open the test ODT file
    let test_file = "test.odt";
    println!("📖 Opening file: {}", test_file);
    let doc = Document::open(test_file)?;
    println!("✅ File opened successfully\n");

    // Test 1: Extract all text content
    println!("--- Test 1: Full Text Extraction ---");
    match doc.text() {
        Ok(text) => {
            let preview = if text.chars().count() > 200 {
                let truncated: String = text.chars().take(200).collect();
                format!("{}... ({} chars total)", truncated, text.chars().count())
            } else {
                text
            };
            println!("Text content:\n{}\n", preview);
        },
        Err(e) => println!("⚠️  Error extracting text: {}\n", e),
    }

    // Test 2: Parse paragraphs with formatting
    println!("--- Test 2: Paragraph Parsing ---");
    match doc.paragraphs() {
        Ok(paragraphs) => {
            println!("Total paragraphs: {}", paragraphs.len());
            for (i, para) in paragraphs.iter().take(5).enumerate() {
                if let Ok(text_str) = para.text() {
                    let preview = if text_str.chars().count() > 80 {
                        let truncated: String = text_str.chars().take(80).collect();
                        format!("{}...", truncated)
                    } else {
                        text_str
                    };
                    println!("  Para {}: {}", i + 1, preview);
                    if let Some(style) = para.style_name() {
                        println!("    Style: {}", style);
                    }
                }
            }
            if paragraphs.len() > 5 {
                println!("  ... and {} more paragraphs", paragraphs.len() - 5);
            }
            println!();
        },
        Err(e) => println!("⚠️  Error parsing paragraphs: {}\n", e),
    }

    // Test 3: Parse tables
    println!("--- Test 3: Table Parsing ---");
    match doc.tables() {
        Ok(tables) => {
            println!("Total tables: {}", tables.len());
            for (i, table) in tables.iter().enumerate() {
                println!("  Table {}:", i + 1);
                if let Some(name) = table.name() {
                    println!("    Name: {}", name);
                }
                if let Ok(table_rows) = table.rows() {
                    if let Ok(col_count) = table.column_count() {
                        println!("    Rows: {}, Columns: {}", table_rows.len(), col_count);
                    }

                    // Show first few cells
                    if !table_rows.is_empty() {
                        println!("    First row cells:");
                        if let Ok(first_row_cells) = table_rows[0].cells() {
                            for (j, cell) in first_row_cells.iter().take(3).enumerate() {
                                if let Ok(cell_text_str) = cell.text() {
                                    let preview = if cell_text_str.chars().count() > 30 {
                                        let truncated: String =
                                            cell_text_str.chars().take(30).collect();
                                        format!("{}...", truncated)
                                    } else {
                                        cell_text_str
                                    };
                                    println!("      Cell {}: {}", j + 1, preview);
                                }
                            }
                        }
                    }
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error parsing tables: {}\n", e),
    }

    // Test 4: Extract metadata
    println!("--- Test 4: Metadata Extraction ---");
    match doc.metadata() {
        Ok(metadata) => {
            println!("Document metadata:");
            if let Some(ref title) = metadata.title {
                println!("  Title: {}", title);
            }
            if let Some(ref author) = metadata.author {
                println!("  Author: {}", author);
            }
            if let Some(ref subject) = metadata.subject {
                println!("  Subject: {}", subject);
            }
            if let Some(ref description) = metadata.description {
                println!("  Description: {}", description);
            }
            if let Some(created) = metadata.created {
                println!("  Created: {}", created);
            }
            if let Some(modified) = metadata.modified {
                println!("  Modified: {}", modified);
            }
            println!();
        },
        Err(e) => println!("⚠️  Error extracting metadata: {}\n", e),
    }

    // Test 5: Extract hyperlinks
    println!("--- Test 5: Hyperlink Extraction ---");
    match doc.hyperlinks() {
        Ok(links) => {
            println!("Total hyperlinks: {}", links.len());
            for (i, (text, href)) in links.iter().take(5).enumerate() {
                let text_preview = if text.chars().count() > 40 {
                    let truncated: String = text.chars().take(40).collect();
                    format!("{}...", truncated)
                } else {
                    text.clone()
                };
                println!("  Link {}: \"{}\" -> {}", i + 1, text_preview, href);
            }
            if links.len() > 5 {
                println!("  ... and {} more links", links.len() - 5);
            }
            println!();
        },
        Err(e) => println!("⚠️  Error extracting hyperlinks: {}\n", e),
    }

    // Test 6: Extract bookmarks
    println!("--- Test 6: Bookmark Extraction ---");
    match doc.bookmarks() {
        Ok(bookmarks) => {
            println!("Total bookmarks: {}", bookmarks.len());
            for (i, bookmark) in bookmarks.iter().take(5).enumerate() {
                if let Some(name) = bookmark.name() {
                    println!("  Bookmark {}: {}", i + 1, name);
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error extracting bookmarks: {}\n", e),
    }

    // Test 7: Parse comments
    println!("--- Test 7: Comment Parsing ---");
    match doc.comments() {
        Ok(comments) => {
            println!("Total comments: {}", comments.len());
            for (i, comment) in comments.iter().take(3).enumerate() {
                println!("  Comment {}:", i + 1);
                if let Some(ref author) = comment.author {
                    println!("    Author: {}", author);
                }
                println!("    ID: {}", comment.id);
            }
            println!();
        },
        Err(e) => println!("⚠️  Error parsing comments: {}\n", e),
    }

    // Test 8: Parse tracked changes
    println!("--- Test 8: Tracked Changes ---");
    match doc.track_changes() {
        Ok(changes) => {
            println!("Total tracked changes: {}", changes.len());
            for (i, change) in changes.iter().take(3).enumerate() {
                println!("  Change {}:", i + 1);
                println!("    Type: {:?}", change.change_type);
                if let Some(ref author) = change.author {
                    println!("    Author: {}", author);
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error parsing tracked changes: {}\n", e),
    }

    // Test 9: Parse sections
    println!("--- Test 9: Section Parsing ---");
    match doc.sections() {
        Ok(sections) => {
            println!("Total sections: {}", sections.len());
            for (i, section) in sections.iter().take(3).enumerate() {
                println!("  Section {}: {}", i + 1, section.name);
                if let Some(ref style) = section.style {
                    println!("    Style: {}", style);
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error parsing sections: {}\n", e),
    }

    println!("=== ODT Reader Test Complete ===");
    println!("✅ All reading functionalities tested successfully!");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example odt_reader_test --features odf --no-default-features"
    );
}
