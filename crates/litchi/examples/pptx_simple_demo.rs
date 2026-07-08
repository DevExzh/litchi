//! Simple PPTX demo without media for testing basic features.
//!
//! Run with: cargo run --example pptx_simple_demo

use litchi::ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating simple PPTX demonstration...\n");

    let mut pkg = Package::new()?;
    let pres = pkg.presentation_mut()?;

    // Slide 1: Title
    println!("Creating Slide 1: Title");
    let slide1 = pres.add_slide()?;
    slide1.set_title("Simple Demo");
    slide1.add_text_box(
        "This is a test presentation",
        914400,
        3429000,
        7315200,
        914400,
    );

    // Slide 2: Table
    println!("Creating Slide 2: Table");
    let slide2 = pres.add_slide()?;
    slide2.set_title("Table Test");
    let table_data = vec![
        vec!["A".to_string(), "B".to_string()],
        vec!["1".to_string(), "2".to_string()],
    ];
    slide2.add_table(table_data, 914400, 1828800, 5486400, 1828800);

    // Slide 3: Shapes
    println!("Creating Slide 3: Shapes");
    let slide3 = pres.add_slide()?;
    slide3.set_title("Shape Test");
    slide3.add_rectangle(
        914400,
        1828800,
        1828800,
        1371600,
        Some("FF6B6B".to_string()),
    );
    slide3.add_ellipse(
        3200400,
        1828800,
        1828800,
        1371600,
        Some("4ECDC4".to_string()),
    );

    // Save
    let output_path = "pptx_simple_demo.pptx";
    println!("\nSaving to {}...", output_path);
    pkg.save(output_path)?;
    println!("✓ Done!");

    Ok(())
}
