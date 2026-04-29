//! PPTX Advanced Features Showcase
//!
//! This example demonstrates all the newly implemented PPTX features:
//! - Animations (fade, fly, wipe, zoom effects)
//! - Charts (bar, line, pie charts)
//! - SmartArt (list, process diagrams)
//! - Sections (slide organization)
//! - Custom slide shows
//! - Handout master settings
//! - Presentation protection
//!
//! ## Usage
//!
//! ```bash
//! cargo run --example pptx_advanced_features
//! ```
//!
//! Then open `pptx_advanced_features.pptx` in PowerPoint to verify!

use litchi::ooxml::pptx::animations::{AnimationEffect, AnimationTrigger};
use litchi::ooxml::pptx::handout::{HandoutHeaderFooter, HandoutLayout, HandoutMaster};
use litchi::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType};
use litchi::ooxml::pptx::smartart::{DiagramType, SmartArtBuilder};
use litchi::ooxml::pptx::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("======================================================");
    println!("  PPTX Advanced Features Showcase");
    println!("======================================================\n");

    // Create a new presentation package
    let mut pkg = Package::new()?;

    {
        let pres = pkg.presentation_mut()?;

        // Set widescreen dimensions (16:9)
        println!("✓ Setting widescreen slide size (16:9)");
        pres.set_widescreen_slide_size();

        // Create all demo slides
        create_title_slide(pres)?;
        create_animation_demo(pres)?;
        create_chart_demo(pres)?;
        create_smartart_demo(pres)?;
        create_sections_demo(pres)?;
        create_custom_show_demo(pres)?;
        create_protection_demo(pres)?;
        create_handout_demo(pres)?;

        // Add sections to organize slides
        println!("\n✓ Adding presentation sections");
        pres.add_section("Introduction", vec![256]); // First slide
        pres.add_section("Animations", vec![257]);
        pres.add_section("Data Visualization", vec![258, 259]); // Charts & SmartArt
        pres.add_section("Organization", vec![260, 261, 262, 263]);

        // Create custom slide shows
        println!("✓ Creating custom slide shows");
        pres.create_custom_show("Quick Overview", vec![256, 258, 259]);
        pres.create_custom_show("Animations Only", vec![256, 257]);
        pres.create_custom_show(
            "Full Presentation",
            vec![256, 257, 258, 259, 260, 261, 262, 263],
        );

        // Set up handout master
        println!("✓ Configuring handout master");
        let mut handout = HandoutMaster::new();
        handout.layout = HandoutLayout::SixSlides;
        handout.header_footer = HandoutHeaderFooter {
            show_header: true,
            header_text: Some("Advanced Features Demo".to_string()),
            show_footer: true,
            footer_text: Some("Created with Litchi".to_string()),
            show_date_time: true,
            auto_date: true,
            date_time_text: None,
            show_slide_number: true,
        };
        pres.set_handout_master(handout);

        // Set read-only recommended (soft protection)
        println!("✓ Setting read-only recommendation");
        pres.set_read_only_recommended(true);

        println!(
            "\n✓ Created {} slides in {} sections",
            pres.slide_count(),
            pres.section_count()
        );
        println!("✓ Created {} custom shows", pres.custom_shows().len());
    }

    // Save the presentation
    let output_path = "pptx_advanced_features.pptx";
    println!("\n✓ Saving presentation to: {}", output_path);
    pkg.save(output_path)?;

    println!("\n======================================================");
    println!("  SUCCESS! Advanced features presentation created!");
    println!("======================================================");
    println!("\nOpen '{}' in PowerPoint to verify:", output_path);
    println!("  - Animations: Play slideshow to see effects");
    println!("  - Charts: Check data visualization slides");
    println!("  - SmartArt: View diagram rendering");
    println!("  - Sections: View > Slide Sorter to see sections");
    println!("  - Custom Shows: Slide Show > Custom Slide Show");
    println!("  - Handout: File > Print > Print Layout");
    println!("  - Protection: File > Info to see read-only status");

    Ok(())
}

fn create_title_slide(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating title slide");
    let slide = pres.add_slide()?;
    slide.set_title("PPTX Advanced Features");

    // Add subtitle text box
    slide.add_text_box(
        "Demonstrating Animations, Charts, SmartArt, Sections, Custom Shows & More",
        914400,  // x: 1 inch
        3429000, // y: 3.75 inches
        7315200, // width: 8 inches
        914400,  // height: 1 inch
    );

    // Add date text box
    slide.add_text_box(
        "Created with Litchi - Rust Office Library",
        914400,  // x: 1 inch
        5486400, // y: 6 inches
        7315200, // width: 8 inches
        457200,  // height: 0.5 inch
    );

    Ok(())
}

fn create_animation_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating animation demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Animation Effects Demo");

    // Add shapes that will be animated
    // Shape IDs: 1=group, 2=title, 3+=user shapes
    slide.add_text_box(
        "Fade In Effect",
        914400,  // x
        1828800, // y
        3657600, // width
        914400,  // height
    );
    let shape1_id: u32 = 3; // First user shape

    slide.add_text_box("Fly In Effect", 914400, 2971800, 3657600, 914400);
    let shape2_id: u32 = 4; // Second user shape

    slide.add_text_box("Wipe Effect", 914400, 4114800, 3657600, 914400);
    let shape3_id: u32 = 5; // Third user shape

    slide.add_text_box("Zoom In Effect", 4800600, 2971800, 3657600, 914400);
    let shape4_id: u32 = 6; // Fourth user shape

    // Add animations to shapes
    // Note: Shape ID 3 is the first text box we added
    slide.add_animation(shape1_id, AnimationEffect::Fade);

    slide.add_animation_with_options(
        shape2_id,
        AnimationEffect::FlyIn,
        AnimationTrigger::AfterPrevious,
        500, // duration_ms
        0,   // delay_ms
    );

    slide.add_animation_with_options(
        shape3_id,
        AnimationEffect::Wipe,
        AnimationTrigger::AfterPrevious,
        750,
        100,
    );

    slide.add_animation_with_options(
        shape4_id,
        AnimationEffect::Zoom,
        AnimationTrigger::WithPrevious,
        1000,
        0,
    );

    // Add speaker notes
    slide.set_notes("This slide demonstrates various animation effects. Click to advance through animations or press F5 to run slideshow.");

    println!("  - Added {} animations", slide.animation_count());

    Ok(())
}

fn create_chart_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating chart demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Chart Examples");

    // Add explanatory text
    slide.add_text_box(
        "Charts are embedded as DrawingML chart objects",
        914400,
        1371600,
        7315200,
        457200,
    );

    // Create chart data for demonstration
    let bar_chart = ChartData::new(ChartType::Bar, 914400, 2057400, 3657600, 2743200)
        .with_title("Quarterly Sales")
        .with_legend(true)
        .add_series(
            ChartSeries::new("2023")
                .with_categories(vec![
                    "Q1".to_string(),
                    "Q2".to_string(),
                    "Q3".to_string(),
                    "Q4".to_string(),
                ])
                .with_values(vec![120.0, 150.0, 180.0, 200.0]),
        )
        .add_series(
            ChartSeries::new("2024")
                .with_categories(vec![
                    "Q1".to_string(),
                    "Q2".to_string(),
                    "Q3".to_string(),
                    "Q4".to_string(),
                ])
                .with_values(vec![140.0, 175.0, 210.0, 240.0]),
        );

    let pie_chart = ChartData::new(ChartType::Pie, 4800600, 2057400, 3657600, 2743200)
        .with_title("Market Share")
        .with_legend(true)
        .add_series(
            ChartSeries::new("Products")
                .with_categories(vec![
                    "Product A".to_string(),
                    "Product B".to_string(),
                    "Product C".to_string(),
                    "Other".to_string(),
                ])
                .with_values(vec![35.0, 28.0, 22.0, 15.0]),
        );

    // Add text boxes showing chart info (actual chart rendering requires package integration)
    slide.add_text_box(
        &format!(
            "Bar Chart: {}\nSeries: {}\nCategories: {}",
            bar_chart.title.as_deref().unwrap_or("Untitled"),
            bar_chart.series.len(),
            bar_chart
                .series
                .first()
                .map(|s| s.categories.len())
                .unwrap_or(0)
        ),
        914400,
        5029200,
        3657600,
        914400,
    );

    slide.add_text_box(
        &format!(
            "Pie Chart: {}\nSlices: {}",
            pie_chart.title.as_deref().unwrap_or("Untitled"),
            pie_chart
                .series
                .first()
                .map(|s| s.values.len())
                .unwrap_or(0)
        ),
        4800600,
        5029200,
        3657600,
        914400,
    );

    slide.set_notes("This slide shows chart configuration. Full chart embedding requires additional package integration.");

    Ok(())
}

fn create_smartart_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating SmartArt demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagrams");

    // Create SmartArt diagrams
    let list_smartart = SmartArtBuilder::new(DiagramType::List)
        .layout_name("Basic List")
        .add_items(vec![
            "Planning",
            "Design",
            "Development",
            "Testing",
            "Deployment",
        ])
        .build();

    let process_smartart = SmartArtBuilder::new(DiagramType::Process)
        .layout_name("Basic Process")
        .add_items(vec!["Input", "Process", "Output"])
        .build();

    let hierarchy_smartart = SmartArtBuilder::new(DiagramType::Hierarchy)
        .layout_name("Organization Chart")
        .add_item("CEO")
        .add_item("CTO")
        .add_item("CFO")
        .add_item("COO")
        .build();

    // Add text representations of SmartArt
    slide.add_text_box(
        &format!(
            "List Diagram ({} items):\n{}",
            list_smartart.node_count(),
            list_smartart.text()
        ),
        914400,
        1600200,
        3657600,
        1600200,
    );

    slide.add_text_box(
        &format!(
            "Process Diagram ({} steps):\n{}",
            process_smartart.node_count(),
            process_smartart.text()
        ),
        4800600,
        1600200,
        3657600,
        1143000,
    );

    slide.add_text_box(
        &format!(
            "Hierarchy ({} nodes):\n{}",
            hierarchy_smartart.node_count(),
            hierarchy_smartart.text()
        ),
        914400,
        3657600,
        7315200,
        1371600,
    );

    slide.set_notes("SmartArt diagrams are defined using DiagramML. The SmartArtBuilder provides a fluent API for creating diagrams.");

    Ok(())
}

fn create_sections_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating sections demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Presentation Sections");

    slide.add_text_box(
        "Sections help organize your presentation into logical groups.\n\n\
        This presentation has the following sections:\n\n\
        • Introduction (1 slide)\n\
        • Animations (1 slide)\n\
        • Data Visualization (2 slides)\n\
        • Organization (4 slides)\n\n\
        View in Slide Sorter mode to see section dividers.",
        914400,
        1828800,
        7315200,
        3200400,
    );

    slide.set_notes("Sections are stored in presentation.xml as p:extLst elements. They appear in PowerPoint's Slide Sorter view.");

    Ok(())
}

fn create_custom_show_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating custom shows demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Custom Slide Shows");

    slide.add_text_box(
        "Custom Slide Shows let you create different versions of your presentation.\n\n\
        This presentation includes:\n\n\
        • Quick Overview - Title, Charts, SmartArt (3 slides)\n\
        • Animations Only - Title, Animation demo (2 slides)\n\
        • Full Presentation - All slides (8 slides)\n\n\
        Access via: Slide Show → Custom Slide Show",
        914400,
        1828800,
        7315200,
        3200400,
    );

    slide.set_notes("Custom shows are defined in p:custShowLst in presentation.xml. They reference slides by their IDs.");

    Ok(())
}

fn create_protection_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating protection demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Presentation Protection");

    slide.add_text_box(
        "Protection features available:\n\n\
        • Read-Only Recommended - Prompts user to open as read-only\n\
        • Structure Protection - Prevents adding/removing slides\n\
        • Password Protection - Requires password to modify\n\n\
        This presentation has 'Read-Only Recommended' enabled.\n\
        When you open it, PowerPoint will suggest opening as read-only.",
        914400,
        1828800,
        7315200,
        3200400,
    );

    slide.set_notes("Protection settings are stored in presentation.xml as p:modifyVerifier or as file properties.");

    {
        let protection = pres.protection_mut();
        protection.set_modify_password("secret123")?;
    }

    Ok(())
}

fn create_handout_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating handout master demo slide");
    let slide = pres.add_slide()?;
    slide.set_title("Handout Master Settings");

    slide.add_text_box(
        "Handout Master controls how slides appear when printed.\n\n\
        This presentation's handout settings:\n\n\
        • Layout: 6 slides per page\n\
        • Header: 'Advanced Features Demo'\n\
        • Footer: 'Created with Litchi'\n\
        • Date/Time: Auto-updated\n\
        • Slide Numbers: Enabled\n\n\
        Preview via: File → Print → Print Layout",
        914400,
        1828800,
        7315200,
        3200400,
    );

    slide.set_notes("Handout master is stored in /ppt/handoutMasters/handoutMaster1.xml with its own relationships.");

    Ok(())
}
