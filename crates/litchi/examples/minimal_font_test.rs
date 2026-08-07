use litchi::{docx::Package, fonts::embedding::Mode};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating minimal DOCX with single embedded font...");

    // Create document
    let mut pkg = Package::new()?;

    // Enable font embedding
    pkg.set_font_embedding(Mode::Subset)?;

    // Add simple paragraph with Liberation Sans
    {
        let doc = pkg.document_mut()?;
        let p = doc.add_paragraph();
        let run = p.add_run();
        run.set_text("Test with Liberation Sans font - ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        run.font_name("Liberation Sans");
        run.font_size(24);
    }

    // Save
    pkg.save("minimal_font_test.docx")?;

    println!("✓ Created: minimal_font_test.docx");
    println!();
    println!("Please test in Microsoft Office and report:");
    println!("1. Does the document open without errors?");
    println!("2. Is 'Liberation Sans' displayed or substituted?");
    println!("3. In File > Info > Properties, is font listed as embedded?");
    println!("4. Any error messages or warnings?");

    Ok(())
}
