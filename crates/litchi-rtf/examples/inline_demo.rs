#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Parse the README's inline RTF snippet and inspect the resulting structure.
//!
//! Run from the workspace root:
//!
//! ```bash
//! cargo run -p litchi-rtf --example inline_demo
//! ```

use litchi_rtf::Document;
use litchi_rtf::text::Inline;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // The exact snippet shown in the litchi-rtf README / lib.rs doc-comment.
    let rtf = r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello World!\par}";

    println!("Input RTF source:");
    println!("{rtf}");
    println!("{}", "=".repeat(60));

    let doc = Document::parse(rtf)?;

    println!("Plain text       : {:?}", doc.text());
    println!("Paragraph count  : {}", doc.paragraph_count());
    println!("Text runs        : {}", doc.body().runs().count());
    println!("Fonts in table   : {}", doc.fonts().len());

    // Show each font defined in the document.
    for (i, font) in doc.fonts().iter().enumerate() {
        println!(
            "  font[{}] name={:?} family={:?} charset={:?}",
            i,
            font.name(),
            font.family(),
            font.charset()
        );
    }

    // Traverse semantic paragraphs and inline runs without flattening or
    // cloning the snapshot's retained text.
    println!("\nParagraphs and inlines:");
    println!("{}", "-".repeat(60));
    for (paragraph_index, paragraph) in doc.body().paragraphs().enumerate() {
        println!(
            "Paragraph {} ({} UTF-8 bytes)",
            paragraph_index + 1,
            paragraph.len()
        );
        for inline in paragraph.inlines() {
            match inline {
                Inline::Text(run) => {
                    let format = run.format();
                    println!(
                        "  text={:?} font={:?} bold={} italic={} underline={:?}",
                        run.text(),
                        format.font().map(litchi_rtf::font::Font::name),
                        format.bold(),
                        format.italic(),
                        format.underline()
                    );
                },
                Inline::Break(kind) => println!("  break={kind:?}"),
                _ => println!("  other inline content"),
            }
        }
    }

    Ok(())
}
