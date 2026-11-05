//! Typed custom slide shows with a transactional PresentationML round-trip.

use litchi_pptx::Package;
use litchi_pptx::presentation_properties::metadata::custom_show::{Limits, List, Show};
use litchi_pptx::presentation_properties::metadata::structure;
use std::error::Error as StdError;

const TITLES: [&str; 15] = [
    "Company Overview",
    "Financial Performance",
    "Revenue Breakdown",
    "Cost Analysis",
    "Market Position",
    "Competitive Landscape",
    "Product Roadmap",
    "Technology Stack",
    "Infrastructure",
    "Team Structure",
    "Hiring Plans",
    "Risk Assessment",
    "Mitigation Strategies",
    "Future Projections",
    "Q&A",
];

fn main() -> Result<(), Box<dyn StdError>> {
    println!("=== Custom Slide Shows Example ===\n");

    let mut builder = litchi_pptx::builder::Builder::new()?;
    for title in TITLES {
        builder.presentation()?.add_slide()?.set_title(title);
    }

    // Publish authored slides before the package-aware structure owner reads
    // their stable PresentationML IDs.
    let mut package = builder.build()?;
    let slide_ids = package
        .with_opc(|opc| structure::load(opc, Limits::DEFAULT))?
        .slides
        .into_iter()
        .map(|slide| slide.slide_id)
        .collect::<Vec<_>>();
    assert_eq!(slide_ids.len(), TITLES.len());

    // Build the contextual custom-show model from real package slide IDs.
    let mut shows = List::new();
    shows.create(
        "Executive Summary",
        [0, 4, 9, 14]
            .into_iter()
            .map(|index| slide_ids[index])
            .collect(),
    );
    shows.create(
        "Technical Deep Dive",
        (0..10).map(|index| slide_ids[index]).collect(),
    );
    shows.create(
        "Sales Pitch",
        [0, 4, 9, 10, 11, 14]
            .into_iter()
            .map(|index| slide_ids[index])
            .collect(),
    );
    shows.create(
        "Quick Demo",
        [0, 2, 14]
            .into_iter()
            .map(|index| slide_ids[index])
            .collect(),
    );

    println!("Custom Shows Created: {}", shows.len());
    for show in &shows.shows {
        println!("Show: '{}' (ID: {})", show.name, show.id);
        println!("  Slides: {:?}", show.slide_ids);
        println!("  Slide count: {}", show.slide_count());
    }

    println!("\n--- Model Lookups ---");
    assert_eq!(
        shows
            .get_by_name("Executive Summary")
            .expect("executive show")
            .slide_count(),
        4
    );
    assert_eq!(
        shows
            .get_by_name("Technical Deep Dive")
            .expect("technical show")
            .slide_count(),
        10
    );
    assert!(shows.get_by_name("Nonexistent").is_none());
    assert_eq!(
        shows.get_by_id(0).expect("first show").name,
        "Executive Summary"
    );

    let xml = shows.to_xml(Limits::DEFAULT)?;
    println!("\n--- Generated Typed XML ---");
    println!("XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:custShowLst>"));
    assert!(xml.contains("Executive Summary"));
    assert!(xml.contains("Technical Deep Dive"));

    // The standalone value remains useful for detached composition too.
    let manual = Show::new(100, "Manual Show").with_slides(
        [
            slide_ids[0],
            slide_ids[1],
            slide_ids[2],
            slide_ids[3],
            slide_ids[4],
        ]
        .to_vec(),
    );
    assert_eq!(manual.slide_count(), 5);

    // Publish the typed values atomically into ppt/presentation.xml.
    {
        let mut edit = package.edit_structure()?;
        for show in shows.shows.iter().cloned() {
            edit.add_custom_show(show)?;
        }
        package = edit.commit()?.into_package();
    }

    let graph = package.with_opc(|opc| structure::load(opc, Limits::DEFAULT))?;
    assert_eq!(graph.custom_shows.shows, shows.shows);
    println!(
        "\nPublished {} custom shows transactionally.",
        graph.custom_shows.len()
    );

    // Exercise a second transactional edit and verify it after reopening.
    let quick_demo_id = shows.get_by_name("Quick Demo").expect("quick demo").id;
    {
        let mut edit = package.edit_structure()?;
        assert!(edit.remove_custom_show(quick_demo_id)?);
        package = edit.commit()?.into_package();
    }

    let round_trip = Package::from_bytes(&package.to_bytes()?)?;
    let graph = round_trip.with_opc(|opc| structure::load(opc, Limits::DEFAULT))?;
    assert_eq!(graph.custom_shows.len(), 3);
    assert!(graph.custom_shows.get_by_name("Quick Demo").is_none());
    assert_eq!(
        graph
            .custom_shows
            .get_by_name("Executive Summary")
            .expect("round-tripped executive show")
            .slide_count(),
        4
    );

    println!(
        "Removed 'Quick Demo'; {} shows remain after reopen.",
        graph.custom_shows.len()
    );
    println!("\n✅ Custom slide show round-trip completed successfully!");
    Ok(())
}
