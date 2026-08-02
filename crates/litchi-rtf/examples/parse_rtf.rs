//! Parse an RTF file and print summary information.
//!
//! Run from the workspace root:
//!
//! ```bash
//! cargo run -p litchi-rtf --example parse_rtf
//! cargo run -p litchi-rtf --example parse_rtf -- test-data/rtf/hyperlink.rtf
//! ```

use litchi_rtf::Document;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Take the path from CLI args, or fall back to a small bundled sample.
    let path: PathBuf = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("test-data/rtf/testStyles.rtf"));

    println!("Parsing RTF file: {}", path.display());
    println!("{}", "=".repeat(60));

    let doc = Document::open(&path)?;

    // Document statistics
    let text = doc.text();
    println!("Text length        : {} characters", text.len());
    println!("Paragraph count    : {}", doc.paragraph_count());
    println!("Font table entries : {}", doc.fonts().len());
    println!("Color table entries: {}", doc.colors().len());
    println!("Text runs          : {}", doc.body().runs().count());
    println!("Tables             : {}", doc.tables().len());
    println!("Pictures           : {}", doc.pictures().len());
    println!("Fields             : {}", doc.fields().len());
    println!("Sections           : {}", doc.sections().len());
    println!("Footnotes          : {}", doc.footnotes().count());
    println!("Endnotes           : {}", doc.endnotes().count());
    println!("Revisions          : {}", doc.revisions().len());

    // Document info / metadata
    let info = doc.info();
    println!("\nDocument metadata");
    println!("{}", "-".repeat(60));
    if let Some(title) = info.title.as_ref() {
        println!("Title    : {}", title);
    }
    if let Some(author) = info.author.as_ref() {
        println!("Author   : {}", author);
    }
    if let Some(subject) = info.subject.as_ref() {
        println!("Subject  : {}", subject);
    }
    if let Some(company) = info.company.as_ref() {
        println!("Company  : {}", company);
    }
    if let Some(keywords) = info.keywords.as_ref() {
        println!("Keywords : {}", keywords);
    }
    if let Some(pages) = info.pages {
        println!("Pages    : {}", pages);
    }
    if let Some(words) = info.words {
        println!("Words    : {}", words);
    }
    if info.title.is_none()
        && info.author.is_none()
        && info.subject.is_none()
        && info.company.is_none()
        && info.keywords.is_none()
    {
        println!("(no metadata fields populated)");
    }

    // Show a small text preview, capped to 400 chars to keep output tidy.
    println!("\nText preview");
    println!("{}", "-".repeat(60));
    let preview: String = text.chars().take(400).collect();
    println!("{}", preview);
    if text.chars().count() > 400 {
        println!("...");
    }

    Ok(())
}
