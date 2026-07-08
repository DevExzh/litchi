//! Custom slide shows example - demonstrates named slide subsets.

use litchi::ooxml::pptx::{CustomShow, CustomShowList};

fn main() {
    println!("=== Custom Slide Shows Example ===\n");

    // Create a custom show list
    let mut shows = CustomShowList::new();

    // Create shows for different audiences
    shows.create("Executive Summary", vec![1, 5, 10, 15]);
    shows.create("Technical Deep Dive", vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    shows.create("Sales Pitch", vec![1, 5, 10, 11, 12, 15]);
    shows.create("Quick Demo", vec![1, 3, 15]);

    println!("Custom Shows Created: {}", shows.len());
    println!();

    // List all shows
    for show in &shows.shows {
        println!("Show: '{}' (ID: {})", show.name, show.id);
        println!("  Slides: {:?}", show.slide_ids);
        println!("  Slide count: {}", show.slide_count());
    }

    // Lookup by name
    println!("\n--- Lookup by Name ---");
    if let Some(exec) = shows.get_by_name("Executive Summary") {
        println!("Found 'Executive Summary': {} slides", exec.slide_count());
    }

    if let Some(tech) = shows.get_by_name("Technical Deep Dive") {
        println!("Found 'Technical Deep Dive': {} slides", tech.slide_count());
    }

    assert!(shows.get_by_name("Nonexistent").is_none());
    println!("'Nonexistent' not found (as expected)");

    // Lookup by ID
    println!("\n--- Lookup by ID ---");
    if let Some(show) = shows.get_by_id(0) {
        println!("ID 0: '{}'", show.name);
    }
    if let Some(show) = shows.get_by_id(2) {
        println!("ID 2: '{}'", show.name);
    }

    // Generate XML
    let xml = shows.to_xml();
    println!("\n--- Generated XML ---");
    println!("XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:custShowLst>"));
    assert!(xml.contains("Executive Summary"));
    assert!(xml.contains("Technical Deep Dive"));

    // Create a show manually
    println!("\n--- Manual Show Creation ---");
    let mut manual_show = CustomShow::new(100, "Manual Show");
    manual_show.add_slide(1);
    manual_show.add_slide(2);
    manual_show.add_slides(vec![3, 4, 5]);

    println!("Manual show: {} slides", manual_show.slide_count());
    assert_eq!(manual_show.slide_ids, vec![1, 2, 3, 4, 5]);

    // Remove a show
    println!("\n--- Remove Show ---");
    let removed = shows.remove_by_name("Quick Demo");
    println!("Removed: {:?}", removed.map(|s| s.name));
    println!("Remaining shows: {}", shows.len());
    assert_eq!(shows.len(), 3);

    println!("\n✅ Custom slide shows example completed successfully!");
}
