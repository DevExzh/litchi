use litchi::ooxml::docx::Package;
use std::error::Error;

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    println!("=== DOCX Package Round-trip Test ===\n");

    let output_doc = "test_output_from_package.docx";

    println!("Test 1: Creating a new DOCX package...");
    let mut pkg = Package::new()?;
    {
        let doc = pkg.document_mut()?;
        doc.add_heading("DOCX Package API Demo", 1)?;
        doc.add_paragraph_with_text(
            "This example exercises the currently supported DOCX package and mutable document APIs.",
        );
        doc.add_heading("Features", 2)?;
        doc.add_paragraph_with_text(
            "It creates a new package, writes paragraphs, saves the document, and reopens it for read-back verification.",
        );
    }
    println!("✓ Created package and populated document content");

    println!("\nTest 2: Saving package...");
    pkg.save(output_doc)?;
    println!("✓ Saved DOCX file: {}", output_doc);

    println!("\nTest 3: Reopening saved package...");
    let reopened = Package::open(output_doc)?;
    let doc = reopened.document()?;
    let paragraphs = doc.paragraphs()?;
    println!("  Paragraph count: {}", paragraphs.len());

    for (i, para) in paragraphs.iter().take(3).enumerate() {
        let text = para.text()?;
        if !text.trim().is_empty() {
            println!("  Paragraph {}: {:?}", i + 1, text);
        }
    }
    println!("✓ Successfully reopened and read back document content");

    println!("\n=== Test Summary ===");
    println!("✓ DOCX package creation works");
    println!("✓ Mutable document editing works");
    println!("✓ DOCX save/open round-trip works");
    println!("✓ Paragraph extraction works");
    println!("Generated file: {}", output_doc);

    Ok(())
}
