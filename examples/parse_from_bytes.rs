/// Example demonstrating parsing Office files from byte buffers.
///
/// This is useful for:
/// - Parsing files from network traffic without creating temporary files
/// - Real-time file analysis from streams
/// - Memory-efficient processing of in-memory data
/// - Integration with web services and APIs

use litchi::{Document, Presentation};
use std::fs;
use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Litchi from_bytes() API Example ===\n");
    println!("Demonstrating zero-copy, memory-efficient parsing\n");

    // Example 1: Parse Word document from bytes
    println!("1. Parsing Word Document from Memory");
    println!("   Simulating network data reception...\n");

    // Simulate receiving file data from network
    let doc_bytes = fs::read("test.doc")?;
    println!("   Received: {} bytes", doc_bytes.len());

    let start = Instant::now();
    
    // Parse directly from bytes - no temporary file needed!
    let doc = Document::from_bytes(doc_bytes)?;
    
    let parse_time = start.elapsed();
    println!("   Parsed in: {:?}", parse_time);
    println!("   Paragraphs: {}", doc.paragraph_count()?);
    
    let text = doc.text()?;
    println!("   Text preview: {}\n", text.chars().take(100).collect::<String>());

    // Example 2: Parse PowerPoint from bytes
    println!("2. Parsing PowerPoint Presentation from Memory");
    println!("   Simulating real-time stream processing...\n");

    let ppt_bytes = fs::read("test.ppt")?;
    println!("   Stream data: {} bytes", ppt_bytes.len());

    let start = Instant::now();
    
    // Parse directly from bytes
    let pres = Presentation::from_bytes(ppt_bytes)?;
    
    let parse_time = start.elapsed();
    println!("   Parsed in: {:?}", parse_time);
    println!("   Slides: {}", pres.slide_count()?);

    // Example 3: Efficient memory usage
    println!("\n3. Efficient Memory Usage");
    println!("   Parsing and extracting data efficiently...\n");

    let data = fs::read("test.docx")?;
    println!("   Data size: {} bytes", data.len());

    let start = Instant::now();
    
    // Parse from bytes - takes ownership but efficient
    let doc = Document::from_bytes(data)?;
    
    let parse_time = start.elapsed();
    println!("   Parsed in: {:?}", parse_time);
    
    // Extract what we need
    let para_count = doc.paragraph_count()?;
    let text_preview = doc.text()?.chars().take(150).collect::<String>();
    
    println!("   Paragraphs: {}", para_count);
    println!("   Preview: {}\n", text_preview);

    // Example 4: Real-time parsing simulation
    println!("4. Real-Time Parsing Simulation");
    println!("   Processing multiple files in sequence...\n");

    let files = vec![
        ("test.doc", "Word (OLE2)"),
        ("test.docx", "Word (OOXML)"),
        ("test.ppt", "PowerPoint (OLE2)"),
        ("test.pptx", "PowerPoint (OOXML)"),
    ];

    let mut total_bytes = 0u64;
    let overall_start = Instant::now();

    for (filename, format) in files {
        if let Ok(bytes) = fs::read(filename) {
            total_bytes += bytes.len() as u64;
            let file_start = Instant::now();

            match filename.split('.').last().unwrap() {
                "doc" | "docx" => {
                    let doc = Document::from_bytes(bytes)?;
                    let count = doc.paragraph_count()?;
                    println!("   ✓ {} ({}) - {} paragraphs - {:?}",
                        filename, format, count, file_start.elapsed());
                }
                "ppt" | "pptx" => {
                    let pres = Presentation::from_bytes(bytes)?;
                    let count = pres.slide_count()?;
                    println!("   ✓ {} ({}) - {} slides - {:?}",
                        filename, format, count, file_start.elapsed());
                }
                _ => {}
            }
        }
    }

    let overall_time = overall_start.elapsed();
    let throughput = (total_bytes as f64) / overall_time.as_secs_f64() / 1_000_000.0;

    println!("\n   Total processed: {} bytes", total_bytes);
    println!("   Total time: {:?}", overall_time);
    println!("   Throughput: {:.2} MB/s", throughput);

    // Example 5: Performance comparison
    println!("\n5. Performance Insights");
    println!("   ───────────────────────────────────────");
    println!("   ✓ No temporary files created");
    println!("   ✓ Direct memory parsing");
    println!("   ✓ Automatic format detection");
    println!("   ✓ Minimal allocations");
    println!("   ✓ Suitable for long-running services");
    println!("   ✓ Safe for concurrent processing");

    println!("\n=== Example Complete ===");
    println!("\nUse Cases:");
    println!("  • Web services receiving Office files");
    println!("  • Network traffic analysis");
    println!("  • Document processing pipelines");
    println!("  • Real-time file monitoring");
    println!("  • Cloud storage integrations");
    
    Ok(())
}

