//! Comprehensive typed custom-show and section example.
//!
//! The slide writer owns slide content, while the presentation-structure
//! owner transactionally publishes custom shows and sections into the OPC
//! presentation graph.

use litchi_pptx::presentation_properties::metadata::custom_show::Show;
use litchi_pptx::presentation_properties::metadata::sections::Section;
use litchi_pptx::presentation_properties::metadata::structure;
use litchi_pptx::{MutablePresentation, Package};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_828_800;
const WIDTH: i64 = 7_315_200;

struct ShowPlan {
    name: &'static str,
    positions: Vec<usize>,
}

struct SectionPlan {
    name: &'static str,
    id: &'static str,
    positions: Vec<usize>,
}

fn main() -> Result<(), Box<dyn StdError>> {
    println!("=== Comprehensive Custom Slide Show Tests ===\n");

    test_executive_presentation()?;
    test_training_presentation()?;
    test_product_demo()?;
    test_conference_presentation()?;

    println!("\n=== All custom show tests complete! ===");
    println!("\nEach output contains typed custom shows in ppt/presentation.xml.");
    Ok(())
}

fn add_slides(
    presentation: &mut MutablePresentation,
    titles: &[&str],
    content_prefix: &str,
    height: i64,
) -> Result<(), Box<dyn StdError>> {
    for title in titles {
        let slide = presentation.add_slide()?;
        slide.set_title(title);
        slide.add_text_box(&format!("{content_prefix}{title}"), X, Y, WIDTH, height);
    }
    Ok(())
}

fn publish_structure(
    mut package: Package,
    output: &str,
    slide_count: usize,
    shows: Vec<ShowPlan>,
    sections: Vec<SectionPlan>,
) -> Result<(), Box<dyn StdError>> {
    // Publish the mutable slide snapshot first. The structure owner operates
    // on the canonical OPC graph and therefore cannot observe stale writer
    // state.
    let bytes = package.to_bytes()?;
    let mut package = Package::from_bytes(&bytes)?;
    let slide_ids = package
        .with_opc(structure::load)?
        .slides
        .into_iter()
        .map(|slide| slide.slide_id)
        .collect::<Vec<_>>();
    assert_eq!(slide_ids.len(), slide_count);

    package.edit_opc(|opc| {
        for (id, plan) in shows.iter().enumerate() {
            let selected = plan
                .positions
                .iter()
                .map(|&position| slide_ids[position])
                .collect::<Vec<_>>();
            structure::add_custom_show(
                opc,
                Show::new((id + 1) as u32, plan.name).with_slides(selected),
            )?;
        }

        for plan in &sections {
            let selected = plan
                .positions
                .iter()
                .map(|&position| slide_ids[position])
                .collect::<Vec<_>>();
            structure::add_section(opc, Section::new(plan.name, plan.id).with_slides(selected))?;
        }
        Ok(())
    })?;

    let graph = package.with_opc(structure::load)?;
    assert_eq!(graph.slides.len(), slide_count);
    assert_eq!(graph.custom_shows.shows.len(), shows.len());
    assert_eq!(graph.sections.len(), sections.len());
    for (actual, expected) in graph.custom_shows.shows.iter().zip(&shows) {
        assert_eq!(actual.name, expected.name);
        assert_eq!(actual.slide_ids.len(), expected.positions.len());
    }

    package.save(output)?;
    let reopened = Package::open(output)?;
    let graph = reopened.with_opc(structure::load)?;
    assert_eq!(graph.custom_shows.shows.len(), shows.len());
    assert_eq!(graph.sections.len(), sections.len());
    println!(
        "  ✓ Saved and reopened {output} ({} slides, {} custom shows, {} sections)",
        graph.slides.len(),
        graph.custom_shows.shows.len(),
        graph.sections.len()
    );
    Ok(())
}

fn test_executive_presentation() -> Result<(), Box<dyn StdError>> {
    println!("Test 1: Creating executive presentation with custom shows...");
    let titles = [
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
    let mut package = Package::new()?;
    add_slides(
        package.presentation_mut()?,
        &titles,
        "Executive content: ",
        2_500_000,
    )?;
    publish_structure(
        package,
        "customshow_executive.pptx",
        titles.len(),
        vec![
            ShowPlan {
                name: "Board Meeting",
                positions: vec![0, 1, 4, 13, 14],
            },
            ShowPlan {
                name: "Executive Summary",
                positions: vec![0, 1, 2, 3, 4, 5, 6, 13, 14],
            },
            ShowPlan {
                name: "Financial Review",
                positions: vec![0, 1, 2, 3, 11, 12, 13, 14],
            },
            ShowPlan {
                name: "Technical Review",
                positions: vec![0, 6, 7, 8, 9, 14],
            },
            ShowPlan {
                name: "All Hands Meeting",
                positions: (0..titles.len()).collect(),
            },
        ],
        vec![
            SectionPlan {
                name: "Overview",
                id: "{11111111-1111-1111-1111-111111111111}",
                positions: vec![0, 1],
            },
            SectionPlan {
                name: "Financials",
                id: "{22222222-2222-2222-2222-222222222222}",
                positions: vec![2, 3],
            },
            SectionPlan {
                name: "Market",
                id: "{33333333-3333-3333-3333-333333333333}",
                positions: vec![4, 5],
            },
            SectionPlan {
                name: "Product & Tech",
                id: "{44444444-4444-4444-4444-444444444444}",
                positions: vec![6, 7, 8, 9],
            },
            SectionPlan {
                name: "Organization",
                id: "{55555555-5555-5555-5555-555555555555}",
                positions: vec![10, 11],
            },
            SectionPlan {
                name: "Strategy",
                id: "{66666666-6666-6666-6666-666666666666}",
                positions: vec![12, 13],
            },
            SectionPlan {
                name: "Closing",
                id: "{77777777-7777-7777-7777-777777777777}",
                positions: vec![14],
            },
        ],
    )
}

fn test_training_presentation() -> Result<(), Box<dyn StdError>> {
    println!("Test 2: Creating training presentation with role-specific shows...");
    let titles = [
        "Welcome & Agenda",
        "Company Policies",
        "Safety Guidelines",
        "Admin: Manager Tools",
        "Admin: Reporting",
        "Admin: Approvals",
        "HR: Employee Records",
        "HR: Benefits Admin",
        "HR: Performance Reviews",
        "IT: System Access",
        "IT: Security Protocols",
        "IT: Troubleshooting",
        "Sales: CRM Training",
        "Sales: Pricing Tools",
        "Sales: Proposal Process",
        "Support: Ticket System",
        "Support: Escalation",
        "Support: Knowledge Base",
        "Q&A Session",
        "Feedback & Next Steps",
    ];
    let mut package = Package::new()?;
    add_slides(
        package.presentation_mut()?,
        &titles,
        "Training content for ",
        3_000_000,
    )?;
    publish_structure(
        package,
        "customshow_training.pptx",
        titles.len(),
        vec![
            ShowPlan {
                name: "New Employee Orientation",
                positions: vec![0, 1, 2, 19],
            },
            ShowPlan {
                name: "Manager Training",
                positions: vec![0, 1, 2, 3, 4, 5, 19],
            },
            ShowPlan {
                name: "HR Department",
                positions: vec![0, 1, 2, 6, 7, 8, 19],
            },
            ShowPlan {
                name: "IT Department",
                positions: vec![0, 1, 2, 9, 10, 11, 19],
            },
            ShowPlan {
                name: "Sales Team",
                positions: vec![0, 1, 12, 13, 14, 19],
            },
            ShowPlan {
                name: "Support Team",
                positions: vec![0, 1, 15, 16, 17, 19],
            },
        ],
        Vec::new(),
    )
}

fn test_product_demo() -> Result<(), Box<dyn StdError>> {
    println!("Test 3: Creating product demo with feature-specific shows...");
    let titles = [
        "Product Introduction",
        "Core Features Overview",
        "Feature: Analytics Dashboard",
        "Feature: Real-time Reporting",
        "Feature: Data Integration",
        "Feature: Custom Workflows",
        "Feature: API Access",
        "Feature: Mobile App",
        "Feature: Collaboration Tools",
        "Feature: Security & Compliance",
        "Feature: Scalability",
        "Pricing Tiers",
        "Implementation Timeline",
        "Customer Success Stories",
        "ROI Calculator",
        "Next Steps",
    ];
    let mut package = Package::new()?;
    add_slides(
        package.presentation_mut()?,
        &titles,
        "Demo content: ",
        3_500_000,
    )?;
    publish_structure(
        package,
        "customshow_product_demo.pptx",
        titles.len(),
        vec![
            ShowPlan {
                name: "Quick Overview",
                positions: vec![0, 1, 11, 15],
            },
            ShowPlan {
                name: "Analytics Focus",
                positions: vec![0, 1, 2, 3, 4, 11, 13, 14, 15],
            },
            ShowPlan {
                name: "Technical Deep Dive",
                positions: vec![0, 5, 6, 7, 10, 11, 15],
            },
            ShowPlan {
                name: "Business Value",
                positions: vec![0, 1, 11, 13, 14, 15],
            },
            ShowPlan {
                name: "Mobile-First",
                positions: vec![0, 7, 8, 11, 15],
            },
            ShowPlan {
                name: "Full Demo",
                positions: (0..titles.len()).collect(),
            },
        ],
        Vec::new(),
    )
}

fn test_conference_presentation() -> Result<(), Box<dyn StdError>> {
    println!("Test 4: Creating conference presentation with time-based shows...");
    let titles = [
        "Title & Speaker Intro",
        "Problem Statement",
        "Current Solutions & Limitations",
        "Our Approach",
        "Technical Architecture",
        "Implementation Details",
        "Code Example 1",
        "Code Example 2",
        "Performance Benchmarks",
        "Comparison with Alternatives",
        "Real-World Use Cases",
        "Case Study: Company A",
        "Case Study: Company B",
        "Lessons Learned",
        "Future Work",
        "Open Source Release",
        "Community Feedback",
        "Q&A",
        "Thank You & Contact Info",
    ];
    let mut package = Package::new()?;
    add_slides(
        package.presentation_mut()?,
        &titles,
        "Conference content: ",
        3_200_000,
    )?;
    publish_structure(
        package,
        "customshow_conference.pptx",
        titles.len(),
        vec![
            ShowPlan {
                name: "Lightning Talk (5 min)",
                positions: vec![0, 1, 3, 18],
            },
            ShowPlan {
                name: "Short Session (15 min)",
                positions: vec![0, 1, 2, 3, 4, 8, 11, 18],
            },
            ShowPlan {
                name: "Standard Talk (30 min)",
                positions: vec![0, 1, 2, 3, 4, 5, 8, 9, 10, 11, 13, 15, 18],
            },
            ShowPlan {
                name: "Extended Session (45 min)",
                positions: (0..=16).collect(),
            },
            ShowPlan {
                name: "Workshop (90 min)",
                positions: (0..titles.len()).collect(),
            },
        ],
        Vec::new(),
    )
}
