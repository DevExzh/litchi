/// Example: Parse PowerPoint shapes, tables, and images.
///
/// This example demonstrates how to use the litchi library to:
/// - Extract shapes from slides
/// - Access text frames and paragraphs
/// - Read table data
/// - Find pictures/images
/// - Identify placeholders
///
/// Usage:
///   cargo run --example parse_pptx_shapes test.pptx

use litchi::ooxml::pptx::shapes::base::ShapeType;
use litchi::ooxml::pptx::shapes::table::Table;
use litchi::ooxml::pptx::shapes::picture::Picture;
use litchi::ooxml::pptx::shapes::base::Shape as TextShape;
use litchi::ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get the file path from command line arguments
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <file.pptx>", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    println!("Opening PowerPoint presentation: {}", file_path);
    println!("{}", "=".repeat(70));

    // Open the package
    let pkg = Package::open(file_path)?;
    let pres = pkg.presentation()?;

    // Get all slides
    let slides = pres.slides()?;
    
    println!("\n📊 Presentation has {} slides\n", slides.len());

    // Process each slide
    for (slide_idx, slide) in slides.iter().enumerate() {
        println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        println!("🎯 SLIDE #{}: {}", 
            slide_idx + 1, 
            slide.name().unwrap_or_else(|_| "(unnamed)".to_string())
        );
        println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

        // Get shapes
        let shapes = slide.shapes()?;
        println!("\n  📦 Total shapes: {}", shapes.len());

        if shapes.is_empty() {
            println!("     (No shapes on this slide)\n");
            continue;
        }

        // Count shapes by type
        let mut text_shapes = 0;
        let mut pictures = 0;
        let mut tables = 0;
        let mut others = 0;

        for shape in &shapes {
            match shape.shape_type() {
                ShapeType::Shape => text_shapes += 1,
                ShapeType::Picture => pictures += 1,
                ShapeType::GraphicFrame => tables += 1,
                _ => others += 1,
            }
        }

        println!("     • Text shapes: {}", text_shapes);
        println!("     • Pictures: {}", pictures);
        println!("     • Tables: {}", tables);
        if others > 0 {
            println!("     • Other: {}", others);
        }

        // Process each shape
        for (shape_idx, mut shape) in shapes.into_iter().enumerate() {
            println!("\n  ▶ Shape #{}: {} ({})", 
                shape_idx + 1,
                shape.name().unwrap_or_else(|_| "Unnamed".to_string()),
                format!("{:?}", shape.shape_type())
            );

            // Show position and size
            if let (Ok(left), Ok(top), Ok(width), Ok(height)) = 
                (shape.left(), shape.top(), shape.width(), shape.height()) {
                println!("     Position: ({}, {})", left, top);
                println!("     Size: {}x{}", width, height);
            }

            // Check if it's a placeholder
            if shape.is_placeholder() {
                println!("     ✨ This is a placeholder");
            }

            // Process based on shape type
            match shape.shape_type() {
                ShapeType::Shape => {
                    // Text shape - extract text content
                    let text_shape = TextShape::new(shape.xml_bytes().to_vec());
                    
                    if let Ok(text_frame) = text_shape.text_frame() {
                        println!("     📝 Text content:");
                        
                        // Get all text
                        if let Ok(text) = text_frame.text() {
                            if !text.is_empty() {
                                for line in text.lines() {
                                    if !line.trim().is_empty() {
                                        println!("        {}", line.trim());
                                    }
                                }
                            } else {
                                println!("        (empty)");
                            }
                        }

                        // Show paragraph count
                        if let Ok(paras) = text_frame.paragraphs() {
                            println!("     Paragraphs: {}", paras.len());
                        }
                    }
                }
                
                ShapeType::Picture => {
                    // Picture shape - show image info
                    let picture = Picture::new(shape.xml_bytes().to_vec());
                    
                    if let Ok(rid) = picture.image_r_id() {
                        println!("     🖼️  Image relationship ID: {}", rid);
                    }
                }
                
                ShapeType::GraphicFrame if shape.has_table() => {
                    // Table - extract table data
                    println!("     📊 Table content:");
                    
                    if let Ok(table) = Table::from_graphic_frame_xml(shape.xml_bytes()) {
                        if let (Ok(rows), Ok(cols)) = (table.row_count(), table.column_count()) {
                            println!("        Size: {}x{} (rows x columns)", rows, cols);
                            
                            // Extract table data
                            if let Ok(table_rows) = table.rows() {
                                for (row_idx, row) in table_rows.iter().enumerate() {
                                    print!("        Row {}: ", row_idx + 1);
                                    
                                    if let Ok(cells) = row.cells() {
                                        let cell_texts: Vec<String> = cells.iter()
                                            .map(|cell| cell.text().unwrap_or_else(|_| "".to_string()))
                                            .collect();
                                        println!("{}", cell_texts.join(" | "));
                                    } else {
                                        println!();
                                    }
                                }
                            }
                        }
                    }
                }
                
                _ => {
                    // Other shape types
                    println!("     ℹ️  Shape type not fully supported yet");
                }
            }
        }

        println!();
    }

    println!("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    println!("✅ Successfully parsed presentation with shapes!");

    Ok(())
}

