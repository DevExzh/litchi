/// Example demonstrating how to parse and extract information from a .docx file
/// using the litchi OOXML parser.
///
/// This example shows:
/// - Opening an OPC package (Office Open XML)
/// - Accessing the main document part
/// - Iterating over parts and relationships
/// - Efficient XML parsing with zero-copy where possible
use litchi::ooxml::OpcPackage;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Get the file path from command line arguments
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <path-to-docx-file>", args[0]);
        eprintln!("\nExample: {} document.docx", args[0]);
        std::process::exit(1);
    }

    let file_path = &args[1];
    println!("Opening OOXML package: {}", file_path);
    println!("{}", "=".repeat(60));

    // Open the OPC package (uses efficient buffered I/O)
    let package = OpcPackage::open(file_path)?;

    println!("\n📦 Package Information:");
    println!("   Total parts: {}", package.part_count());

    // Display package-level relationships
    println!("\n🔗 Package Relationships:");
    for rel in package.rels().iter() {
        println!(
            "   {} -> {} ({})",
            rel.r_id(),
            rel.target_ref(),
            if rel.is_external() {
                "external"
            } else {
                "internal"
            }
        );
        println!("      Type: {}", rel.reltype());
    }

    // Access the main document part
    println!("\n📄 Main Document Part:");
    match package.main_document_part() {
        Ok(main_part) => {
            println!("   Partname: {}", main_part.partname());
            println!("   Content Type: {}", main_part.content_type());
            println!("   Size: {} bytes", main_part.blob().len());

            // Display part relationships
            let rels_count = main_part.rels().len();
            if rels_count > 0 {
                println!("   Relationships: {}", rels_count);
                for rel in main_part.rels().iter() {
                    println!("      {} -> {}", rel.r_id(), rel.target_ref());
                }
            }
        }
        Err(e) => {
            eprintln!("   Error accessing main document: {}", e);
        }
    }

    // List all parts in the package
    println!("\n📋 All Parts:");
    for (i, part) in package.iter_parts().enumerate() {
        println!("   {}. {}", i + 1, part.partname());
        println!("      Type: {}", part.content_type());
        println!("      Size: {} bytes", part.blob().len());

        // Show extension
        let ext = part.partname().ext();
        if !ext.is_empty() {
            println!("      Extension: .{}", ext);
        }

        println!();
    }

    // Performance statistics
    println!("\n⚡ Performance Notes:");
    println!("   - Uses memchr for fast string searching");
    println!("   - Uses atoi_simd for fast integer parsing");
    println!("   - Uses quick-xml for zero-copy XML parsing");
    println!("   - Minimal allocations through borrowing");

    Ok(())
}
