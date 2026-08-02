//! Parse the README's inline RTF snippet and inspect the resulting structure.
//!
//! Run from the workspace root:
//!
//! ```bash
//! cargo run -p litchi-rtf --example inline_demo
//! ```

use litchi_rtf::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // The exact snippet shown in the litchi-rtf README / lib.rs doc-comment.
    let rtf = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}";

    println!("Input RTF source:");
    println!("{}", rtf);
    println!("{}", "=".repeat(60));

    let doc = Document::parse(rtf)?;

    println!("Plain text       : {:?}", doc.text());
    println!("Paragraph count  : {}", doc.paragraph_count());
    println!("Style blocks     : {}", doc.blocks().len());
    println!("Fonts in table   : {}", doc.fonts().len());

    // Show each font defined in the document.
    for (i, font) in doc.fonts().iter().enumerate() {
        println!(
            "  font[{}] name={:?} family={:?} charset={:?}",
            i, font.name, font.family, font.charset
        );
    }

    // Traverse the parser's borrowed text/formatting blocks without cloning
    // the snapshot or its retained resources.
    println!("\nText blocks:");
    println!("{}", "-".repeat(60));
    for (index, block) in doc.blocks().iter().enumerate() {
        println!(
            "  block[{index}] text={:?} bold={} italic={} underline={:?}",
            block.text, block.formatting.bold, block.formatting.italic, block.formatting.underline
        );
    }

    Ok(())
}
