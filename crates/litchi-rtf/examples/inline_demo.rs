//! Parse the README's inline RTF snippet and inspect the resulting structure.
//!
//! Run from the workspace root:
//!
//! ```bash
//! cargo run -p litchi-rtf --example inline_demo
//! ```

use litchi_rtf::RtfDocument;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // The exact snippet shown in the litchi-rtf README / lib.rs doc-comment.
    let rtf = r"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard Hello World!\par}";

    println!("Input RTF source:");
    println!("{}", rtf);
    println!("{}", "=".repeat(60));

    let doc = RtfDocument::parse(rtf)?;

    println!("Plain text       : {:?}", doc.text());
    println!("Paragraph count  : {}", doc.paragraph_count());
    println!("Style blocks     : {}", doc.blocks().len());
    println!("Fonts in table   : {}", doc.font_table().fonts().len());

    // Show each font defined in the document.
    for (i, font) in doc.font_table().fonts().iter().enumerate() {
        println!(
            "  font[{}] name={:?} family={:?} charset={:?}",
            i, font.name, font.family, font.charset
        );
    }

    // Dump the runs (text + formatting) using `paragraphs_with_content`.
    println!("\nParagraphs with runs:");
    println!("{}", "-".repeat(60));
    for (i, para) in doc.paragraphs_with_content().iter().enumerate() {
        println!("Paragraph {} (text={:?})", i + 1, para.text());
        for (j, run) in para.runs().iter().enumerate() {
            println!(
                "  run[{}] text={:?} bold={:?} italic={:?} underline={}",
                j,
                run.text(),
                run.bold(),
                run.italic(),
                run.underline()
            );
        }
    }

    // `runs()` returns a flat view across the whole document.
    println!("\nTotal runs (flattened): {}", doc.runs().len());

    Ok(())
}
