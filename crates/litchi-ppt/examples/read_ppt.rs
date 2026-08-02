//! Read a legacy PowerPoint `.ppt` file and print slide text.
//!
//! Run with:
//!     cargo run -p litchi-ppt --example read_ppt
//!     cargo run -p litchi-ppt --example read_ppt -- path/to/file.ppt

use litchi_ppt::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "test-data/ole/ppt/SampleShow.ppt".to_string());

    println!("Opening PPT: {}", path);

    let mut package = Package::open(&path)?;
    let pres = package.presentation()?;

    println!("\n=== Presentation ===");
    println!("Slide count : {}", pres.slide_count());
    println!("Has pictures: {}", pres.has_pictures());

    // Whole-presentation text (used as a fallback for slides whose per-slide
    // text extraction returns an empty string).
    let all_text = pres.text().unwrap_or_default();
    let total_chars = all_text.chars().count();
    println!("Total chars : {}", total_chars);

    // Per-slide text.
    let slides = pres.slides()?;
    for slide in &slides {
        let n = slide.slide_number();
        let shape_count = slide.shape_count().unwrap_or(0);
        let text = slide.text().unwrap_or("");
        println!("\n--- Slide {} ({} shapes) ---", n, shape_count);
        if text.trim().is_empty() {
            // Note: per-slide extraction may legitimately be empty for slides
            // that store text only in master/layout records. The presentation
            // -wide `pres.text()` call above gives a fallback aggregated view.
            println!("(no per-slide text extracted)");
        } else {
            // Cap each slide's preview at ~500 chars.
            let preview: String = text.chars().take(500).collect();
            println!("{}", preview);
            let len = text.chars().count();
            if len > 500 {
                println!("... [truncated, {} more chars]", len - 500);
            }
        }
    }

    // If we got no per-slide text at all, dump the aggregated text as a fallback.
    let any_per_slide_text = slides
        .iter()
        .any(|s| s.text().map(|t| !t.trim().is_empty()).unwrap_or(false));
    if !any_per_slide_text && !all_text.trim().is_empty() {
        println!("\n=== Aggregated presentation text (first 500 chars) ===");
        let preview: String = all_text.chars().take(500).collect();
        println!("{}", preview);
    }

    Ok(())
}
