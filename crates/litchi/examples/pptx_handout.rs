//! Handout master example - demonstrates handout configuration.

use litchi::ooxml::pptx::{HandoutLayout, HandoutMaster};
use std::str::FromStr;

fn main() {
    println!("=== Handout Master Example ===\n");

    // Create a basic handout master
    let handout = HandoutMaster::new()
        .with_layout(HandoutLayout::ThreeSlides)
        .with_header("My Presentation Title")
        .with_footer("Confidential - Do Not Distribute")
        .with_slide_numbers()
        .with_date_time();

    println!("Handout Configuration:");
    println!("  Layout: {:?}", handout.layout);
    println!("  Show header: {}", handout.header_footer.show_header);
    println!("  Header text: {:?}", handout.header_footer.header_text);
    println!("  Show footer: {}", handout.header_footer.show_footer);
    println!("  Footer text: {:?}", handout.header_footer.footer_text);
    println!(
        "  Show slide numbers: {}",
        handout.header_footer.show_slide_number
    );
    println!("  Show date/time: {}", handout.header_footer.show_date_time);
    println!("  Auto date: {}", handout.header_footer.auto_date);

    // Generate XML
    let xml = handout.to_xml();
    println!("\nGenerated XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:handoutMaster"));
    assert!(xml.contains("hdr=\"1\""));
    assert!(xml.contains("ftr=\"1\""));

    // Test all handout layouts
    println!("\n--- Handout Layouts ---");
    let layouts = [
        HandoutLayout::OneSlide,
        HandoutLayout::TwoSlides,
        HandoutLayout::ThreeSlides,
        HandoutLayout::FourSlides,
        HandoutLayout::SixSlides,
        HandoutLayout::NineSlides,
        HandoutLayout::Outline,
    ];

    for layout in layouts {
        let str_repr = layout.as_str();
        let parsed = HandoutLayout::from_str(str_repr).unwrap();
        println!("  {:?} -> '{}' -> {:?}", layout, str_repr, parsed);
        assert_eq!(layout, parsed);
    }

    // Create handout with background
    let colored_handout = HandoutMaster::new()
        .with_layout(HandoutLayout::SixSlides)
        .with_background_color("E6E6FA") // Lavender
        .with_footer("Page Footer");

    println!("\nColored Handout:");
    println!("  Layout: {:?}", colored_handout.layout);
    println!("  Background: {:?}", colored_handout.background_color);

    let xml2 = colored_handout.to_xml();
    assert!(xml2.contains("E6E6FA"));

    println!("\n✅ Handout master example completed successfully!");
}
