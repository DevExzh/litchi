//! Comprehensive ODP (OpenDocument Presentation) reading example.
//!
//! This example demonstrates all reading capabilities for ODP files,
//! serving as both verification and regression testing.
//!
//! Run with:
//! ```bash
//! cargo run --example odp_reader_test --features odf --no-default-features
//! ```

#[cfg(feature = "odf")]
use litchi::Result;
#[cfg(feature = "odf")]
use litchi::odf::Presentation;

#[cfg(feature = "odf")]
fn main() -> Result<()> {
    println!("=== ODP Reader Comprehensive Test ===\n");

    // Open the test ODP file
    let test_file = "test.odp";
    println!("📖 Opening file: {}", test_file);
    let presentation = Presentation::open(test_file)?;
    println!("✅ File opened successfully\n");

    // Test 1: Slide count
    println!("--- Test 1: Slide Count ---");
    match presentation.slide_count() {
        Ok(count) => {
            println!("Total slides: {}", count);
            println!();
        },
        Err(e) => println!("⚠️  Error getting slide count: {}\n", e),
    }

    // Test 2: Get all slides
    println!("--- Test 2: Slide Enumeration ---");
    match presentation.slides() {
        Ok(slides) => {
            println!("Retrieved {} slides", slides.len());
            for (i, slide) in slides.iter().enumerate() {
                print!("  Slide {}: ", i + 1);
                if let Ok(Some(title)) = slide.title() {
                    println!("\"{}\"", title);
                } else {
                    println!("(no title)");
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error retrieving slides: {}\n", e),
    }

    // Test 3: Extract text from each slide
    println!("--- Test 3: Text Extraction from Slides ---");
    match presentation.slides() {
        Ok(slides) => {
            for (i, slide) in slides.iter().enumerate() {
                println!("Slide {} text:", i + 1);
                match slide.text() {
                    Ok(text) => {
                        if text.is_empty() {
                            println!("  (No text content)");
                        } else {
                            let preview = if text.chars().count() > 150 {
                                let truncated: String = text.chars().take(150).collect();
                                format!("{}... ({} chars total)", truncated, text.chars().count())
                            } else {
                                text.to_string()
                            };
                            println!("  {}", preview);
                        }
                    },
                    Err(e) => println!("  ⚠️  Error extracting text: {}", e),
                }
                println!();
            }
        },
        Err(e) => println!("⚠️  Error accessing slides: {}\n", e),
    }

    // Test 4: Shape extraction
    println!("--- Test 4: Shape Extraction ---");
    match presentation.slides() {
        Ok(slides) => {
            for (i, slide) in slides.iter().enumerate() {
                println!("Slide {} shapes:", i + 1);
                match slide.shapes() {
                    Ok(shapes) => {
                        if shapes.is_empty() {
                            println!("  (No shapes)");
                        } else {
                            println!("  Total shapes: {}", shapes.len());
                            for (j, shape) in shapes.iter().take(5).enumerate() {
                                println!("    Shape {}: {:?}", j + 1, shape.shape_type());

                                // Extract text from shape if available
                                if let Ok(shape_text) = shape.text()
                                    && !shape_text.is_empty()
                                {
                                    let preview = if shape_text.chars().count() > 60 {
                                        let truncated: String =
                                            shape_text.chars().take(60).collect();
                                        format!("{}...", truncated)
                                    } else {
                                        shape_text.to_string()
                                    };
                                    println!("      Text: {}", preview);
                                }

                                // Show shape properties if available
                                if let Some(ref name) = shape.name {
                                    println!("      Name: {}", name);
                                }
                                if let Some(ref x) = shape.x {
                                    println!("      X: {}", x);
                                }
                                if let Some(ref y) = shape.y {
                                    println!("      Y: {}", y);
                                }
                                if let Some(ref width) = shape.width {
                                    println!("      Width: {}", width);
                                }
                                if let Some(ref height) = shape.height {
                                    println!("      Height: {}", height);
                                }
                            }
                            if shapes.len() > 5 {
                                println!("    ... and {} more shapes", shapes.len() - 5);
                            }
                        }
                    },
                    Err(e) => println!("  ⚠️  Error extracting shapes: {}", e),
                }
                println!();
            }
        },
        Err(e) => println!("⚠️  Error accessing slides: {}\n", e),
    }

    // Test 5: Slide notes
    println!("--- Test 5: Slide Notes ---");
    match presentation.slides() {
        Ok(slides) => {
            let mut notes_found = 0;
            for (i, slide) in slides.iter().enumerate() {
                if let Ok(Some(notes)) = slide.notes() {
                    notes_found += 1;
                    println!("  Slide {} notes: {}", i + 1, notes);
                }
            }
            if notes_found == 0 {
                println!("  (No speaker notes found)");
            }
            println!();
        },
        Err(e) => println!("⚠️  Error accessing slides: {}\n", e),
    }

    // Test 6: Metadata extraction
    println!("--- Test 6: Metadata Extraction ---");
    match presentation.metadata() {
        Ok(metadata) => {
            println!("Presentation metadata:");
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

    // Test 7: Content statistics
    println!("--- Test 7: Content Statistics ---");
    match presentation.slides() {
        Ok(slides) => {
            let mut total_chars = 0;
            let mut slides_with_text = 0;
            let mut slides_with_shapes = 0;
            let mut total_shapes = 0;

            for slide in &slides {
                if let Ok(text) = slide.text()
                    && !text.is_empty()
                {
                    total_chars += text.len();
                    slides_with_text += 1;
                }
                if let Ok(shapes) = slide.shapes()
                    && !shapes.is_empty()
                {
                    slides_with_shapes += 1;
                    total_shapes += shapes.len();
                }
            }

            println!("Presentation statistics:");
            println!("  Total slides: {}", slides.len());
            println!("  Slides with text: {}", slides_with_text);
            println!("  Slides with shapes: {}", slides_with_shapes);
            println!("  Total shapes: {}", total_shapes);
            println!("  Total characters: {}", total_chars);
            println!();
        },
        Err(e) => println!("⚠️  Error computing statistics: {}\n", e),
    }

    // Test 8: Full text extraction
    println!("--- Test 8: Full Text Extraction ---");
    match presentation.text() {
        Ok(text) => {
            let preview = if text.len() > 300 {
                format!("{}... ({} chars total)", &text[..300], text.len())
            } else {
                text
            };
            println!("Full text content:\n{}\n", preview);
        },
        Err(e) => println!("⚠️  Error extracting text: {}\n", e),
    }

    // Test 9: Slide-by-slide summary
    println!("--- Test 9: Complete Slide-by-Slide Summary ---");
    match presentation.slides() {
        Ok(slides) => {
            println!("Complete presentation structure:");
            for (i, slide) in slides.iter().enumerate() {
                println!("\n📊 Slide {} Summary:", i + 1);

                if let Ok(Some(title)) = slide.title() {
                    println!("  Title: {}", title);
                }

                if let Ok(text) = slide.text() {
                    let char_count = text.len();
                    let word_count = text.split_whitespace().count();
                    println!("  Text: {} chars, {} words", char_count, word_count);
                }

                if let Ok(shapes) = slide.shapes() {
                    println!("  Shapes: {}", shapes.len());
                }

                if let Ok(Some(notes)) = slide.notes() {
                    println!("  Notes: {} chars", notes.len());
                }
            }
            println!();
        },
        Err(e) => println!("⚠️  Error generating summary: {}\n", e),
    }

    println!("=== ODP Reader Test Complete ===");
    println!("✅ All reading functionalities tested successfully!");

    Ok(())
}

#[cfg(not(feature = "odf"))]
fn main() {
    eprintln!(
        "This example requires the 'odf' feature. Try: cargo run --example odp_reader_test --features odf --no-default-features"
    );
}
