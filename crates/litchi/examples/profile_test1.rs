//! Profiling example for test_1.docx formula conversion
//!
//! This example is designed to be profiled with samply to identify
//! performance bottlenecks in the formula conversion pipeline.
//!
//! Run with:
//! ```sh
//! cargo build --profile profiling --features all --example profile_test1
//! samply record ./target/profiling/examples/profile_test1
//! ```

use litchi::markdown::{MarkdownOptions, ToMarkdown};

fn main() -> Result<(), litchi::Error> {
    println!("Starting profiling run for test_1.docx...");

    // Open test_1.docx which has 14,937 formulas
    let doc = litchi::open("test_1.docx")?;

    println!("Document loaded. Starting formula conversion...");

    // Configure Markdown options with formula conversion
    let options = MarkdownOptions {
        formula_style: litchi::markdown::FormulaStyle::Dollar,
        include_styles: true,
        ..MarkdownOptions::default()
    };

    // This is the hot path we want to profile
    let _markdown = doc.to_markdown_with_options(&options)?;

    println!("Formula conversion complete!");
    println!("Profile data will be loaded in Firefox Profiler.");

    Ok(())
}
