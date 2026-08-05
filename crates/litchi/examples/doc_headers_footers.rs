//! Example: DOC headers/footers with odd/even/first behavior
//! Run: cargo run --example doc_headers_footers
use litchi_doc::DocWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut w = DocWriter::new();

    // Body text
    w.add_paragraph("This document demonstrates odd/even/first headers and footers.")?;
    w.add_paragraph("Page 1 is a title page with its own header/footer.")?;

    // Headers
    w.set_first_header("First-Page Header");
    w.set_odd_header("Odd Header");
    w.set_even_header("Even Header");

    // Footers
    w.set_first_footer("First-Page Footer");
    w.set_odd_footer("Odd Footer");
    w.set_even_footer("Even Footer");

    // Save
    w.save("headers_footers.doc")?;
    println!("Saved to headers_footers.doc");
    Ok(())
}
