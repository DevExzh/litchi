//! Extract tables from an Apple Numbers (`.numbers`) document and print the
//! first table as CSV.
//!
//! # Run
//!
//! ```bash
//! cargo run -p litchi-iwa --example extract_structured -- /path/to/spreadsheet.numbers
//! ```
//!
//! Numbers test fixtures are not bundled with this checkout. Drop a real
//! `.numbers` file into `test-data/iwa/numbers/` (or anywhere on disk) and
//! pass its path on the command line.

use std::env;
use std::path::Path;

use litchi_iwa::Document;

const TEST_DATA_HINT: &str = "test-data/iwa/numbers";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = match env::args().nth(1) {
        Some(p) => p,
        None => {
            eprintln!("usage: extract_structured <path-to-numbers-file>");
            eprintln!();
            eprintln!(
                "no path given. drop a .numbers file into `{}` and pass its path,",
                TEST_DATA_HINT
            );
            eprintln!("or point at any Numbers document on disk.");
            return Ok(());
        },
    };

    let path = Path::new(&path);
    if !path.exists() {
        eprintln!("file not found: {}", path.display());
        eprintln!("Numbers test fixtures are not committed; please supply a real .numbers path.");
        return Ok(());
    }

    println!("opening: {}", path.display());
    let doc = Document::open(path)?;
    let structured = doc.extract_structured_data()?;

    println!(
        "found {} table(s), {} slide(s), {} section(s)",
        structured.tables.len(),
        structured.slides.len(),
        structured.sections.len()
    );

    let Some(first) = structured.tables.first() else {
        println!("no tables found in this document.");
        return Ok(());
    };

    println!(
        "--- table: {} ({} rows x {} cols) ---",
        first.name(),
        first.row_count(),
        first.column_count()
    );
    let csv = first.to_csv();
    if csv.is_empty() {
        println!("(table is empty)");
    } else {
        println!("{}", csv);
    }

    if structured.tables.len() > 1 {
        println!(
            "... and {} more table(s) not shown.",
            structured.tables.len() - 1
        );
    }

    Ok(())
}
