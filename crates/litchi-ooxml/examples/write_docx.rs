//! Build a fresh `.docx` (heading + bold/italic paragraph + small table),
//! save it to a tempfile, reopen it, and print the round-tripped text.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-ooxml --example write_docx --all-features
//! ```

use litchi_ooxml::docx::{Package, ParagraphAlignment};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Building a new DOCX in memory...");

    let mut pkg = Package::new()?;

    // Set core properties.
    {
        let props = pkg.properties_mut();
        props.title = Some("litchi-ooxml example".to_string());
        props.creator = Some("litchi-ooxml write_docx example".to_string());
        props.subject = Some("Round-trip demo".to_string());
    }

    // Build the body of the document.
    {
        let doc = pkg.document_mut()?;

        // Heading 1
        doc.add_heading("Round-Trip Demo", 1)?;

        // Paragraph mixing bold and italic runs.
        let para = doc.add_paragraph();
        para.set_alignment(ParagraphAlignment::Left);
        para.add_run_with_text("This paragraph contains ");
        para.add_run_with_text("bold").bold(true);
        para.add_run_with_text(", ");
        para.add_run_with_text("italic").italic(true);
        para.add_run_with_text(", and ");
        para.add_run_with_text("bold-italic")
            .bold(true)
            .italic(true);
        para.add_run_with_text(" runs.");

        // Small 2x2 table with a header row.
        let table = doc.add_table(2, 2);
        table.set_width_percent(60);
        if let Some(c) = table.cell(0, 0) {
            let p = c.add_paragraph();
            p.add_run_with_text("Name").bold(true);
        }
        if let Some(c) = table.cell(0, 1) {
            let p = c.add_paragraph();
            p.add_run_with_text("Value").bold(true);
        }
        if let Some(c) = table.cell(1, 0) {
            c.set_text("answer");
        }
        if let Some(c) = table.cell(1, 1) {
            c.set_text("42");
        }
    }

    // Save to a tempfile (with .docx extension so reopen path-detection is unambiguous).
    let tmp = NamedTempFile::with_suffix(".docx")?;
    let tmp_path = tmp.path().to_path_buf();
    println!("Saving to {}", tmp_path.display());
    pkg.save(&tmp_path)?;

    // Reopen and verify.
    println!("Reopening to verify round-trip...");
    let reopened = Package::open(&tmp_path)?;
    let doc = reopened.document()?;

    let paragraph_count = doc.paragraph_count()?;
    let table_count = doc.table_count()?;
    let text = doc.text()?;

    println!("Round-tripped paragraph count: {}", paragraph_count);
    println!("Round-tripped table count    : {}", table_count);
    println!("\n--- Round-tripped text ---\n{}", text);

    let props = reopened.properties();
    println!("\nRound-tripped title   : {:?}", props.title);
    println!("Round-tripped creator : {:?}", props.creator);
    println!("Round-tripped subject : {:?}", props.subject);

    // tmp is dropped here, removing the file.
    Ok(())
}
