/// Example: Parse a PowerPoint presentation and extract information.
///
/// This example demonstrates how to use the litchi library to:
/// - Open a PowerPoint presentation
/// - Get presentation metadata (dimensions, slide count)
/// - Extract text from slides
/// - Access slide masters
///
/// Usage:
///   cargo run --example parse_pptx test.pptx

use litchi::ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get the file path from command line arguments
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} [file.pptx]", args[0]);
        eprintln!("\nExamples:");
        eprintln!("  cargo run --example parse_pptx -- presentation.pptx");
        eprintln!("  cargo run --example parse_pptx    # uses test.pptx");
        std::process::exit(1);
    }

    let file_path = if args.len() > 1 {
        &args[1]
    } else {
        "test.pptx"
    };

    println!("Opening PowerPoint presentation: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the package
    let pkg = Package::open(file_path)?;
    let pres = pkg.presentation()?;

    // Print presentation info
    println!("\n📊 Presentation Info:");
    println!("{}", "-".repeat(60));
    
    let slide_count = pres.slide_count()?;
    println!("  Slide count: {}", slide_count);

    if let (Some(width), Some(height)) = (pres.slide_width()?, pres.slide_height()?) {
        // Convert EMUs to inches (1 inch = 914400 EMUs)
        let width_inches = width as f64 / 914400.0;
        let height_inches = height as f64 / 914400.0;
        println!("  Slide dimensions:");
        println!("    Width:  {} EMUs ({:.2} inches, {:.2} cm)", width, width_inches, width_inches * 2.54);
        println!("    Height: {} EMUs ({:.2} inches, {:.2} cm)", height, height_inches, height_inches * 2.54);
    } else {
        println!("  Slide dimensions: Not defined");
    }

    // Extract and print slide content
    println!("\n📝 Slides:");
    println!("{}", "-".repeat(60));

    let slides = pres.slides()?;

    if slides.is_empty() {
        println!("  No slides found");
    } else {
        for (idx, slide) in slides.iter().enumerate() {
            println!("\n  Slide #{} - {}", idx + 1, slide.name().unwrap_or_else(|_| "(unnamed)".to_string()));

            // Extract text content
            let text = slide.text()?;
            if !text.is_empty() {
                println!("  Text content:");
                let lines: Vec<&str> = text.lines().collect();
                println!("    Total lines: {}", lines.len());

                // Show first 15 lines of content
                for (line_idx, line) in lines.iter().take(15).enumerate() {
                    if !line.trim().is_empty() {
                        println!("    {:2}: {}", line_idx + 1, line);
                    }
                }

                if lines.len() > 15 {
                    println!("    ... ({} more lines)", lines.len() - 15);
                }
            } else {
                println!("    (No text content)");
            }

            // Show shapes information
            match slide.shapes() {
                Ok(shapes) => {
                    println!("    Shapes: {} total", shapes.len());
                    // Show shape types
                    let mut shape_types = std::collections::HashMap::new();
                    for shape in &shapes {
                        *shape_types.entry(format!("{:?}", shape.shape_type())).or_insert(0) += 1;
                    }
                    if !shape_types.is_empty() {
                        println!("      Types: {:?}", shape_types);
                    }
                }
                Err(_) => {
                    println!("    Shapes: (Error accessing)");
                }
            }
        }
    }

    // Print slide master info
    println!("\n🎨 Slide Masters:");
    println!("{}", "-".repeat(60));
    
    let masters = pres.slide_masters()?;
    
    if masters.is_empty() {
        println!("  No slide masters found");
    } else {
        for (idx, master) in masters.iter().enumerate() {
            let name = master.name().unwrap_or_else(|_| "(unnamed)".to_string());
            let layout_count = master.slide_layout_rids().map(|rids| rids.len()).unwrap_or(0);
            println!("  Master #{}: {} ({} layouts)", idx + 1, name, layout_count);
        }
    }

    println!("\n{}", "=".repeat(60));
    println!("✅ Successfully parsed presentation!");

    Ok(())
}

