//! Comprehensive Custom Slide Show Example
//!
//! Demonstrates creating presentations with custom slide shows for different audiences.
//! Custom shows allow you to present subsets of slides without creating multiple files.

use litchi::ooxml::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive Custom Slide Show Tests ===\n");

    // Test 1: Executive presentation with multiple custom shows
    test_executive_presentation()?;

    // Test 2: Training presentation with role-specific shows
    test_training_presentation()?;

    // Test 3: Product demo with feature-specific shows
    test_product_demo()?;

    // Test 4: Conference presentation with time-based shows
    test_conference_presentation()?;

    println!("\n=== All custom show tests complete! ===");
    println!("\nTo verify custom shows:");
    println!("  1. Open each file in PowerPoint");
    println!("  2. Go to Slide Show > Custom Slide Show");
    println!("  3. Verify that custom shows appear in the list");
    println!("  4. Run each custom show to verify slide selection");

    Ok(())
}

/// Test 1: Executive Presentation
/// Different custom shows for different management levels
fn test_executive_presentation() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 1: Creating executive presentation with custom shows...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Create 15 slides with content
        let slide_titles = vec![
            "Company Overview",      // 1
            "Financial Performance", // 2
            "Revenue Breakdown",     // 3
            "Cost Analysis",         // 4
            "Market Position",       // 5
            "Competitive Landscape", // 6
            "Product Roadmap",       // 7
            "Technology Stack",      // 8
            "Infrastructure",        // 9
            "Team Structure",        // 10
            "Hiring Plans",          // 11
            "Risk Assessment",       // 12
            "Mitigation Strategies", // 13
            "Future Projections",    // 14
            "Q&A",                   // 15
        ];

        for (i, title) in slide_titles.iter().enumerate() {
            let slide = pres.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(
                &format!("Slide {} content area", i + 1),
                914400,
                1828800,
                7315200,
                2500000,
            );
        }

        // Add sections for organization
        pres.add_section("Overview", vec![256, 257]);
        pres.add_section("Financials", vec![258, 259]);
        pres.add_section("Market", vec![260, 261]);
        pres.add_section("Product & Tech", vec![262, 263, 264, 265]);
        pres.add_section("Organization", vec![266, 267]);
        pres.add_section("Strategy", vec![268, 269]);
        pres.add_section("Closing", vec![270]);

        // Create custom shows for different audiences
        pres.create_custom_show("Board Meeting", vec![256, 257, 260, 269, 270]);
        println!("  ✓ Board Meeting show: 5 slides");

        pres.create_custom_show(
            "Executive Summary",
            vec![256, 257, 258, 259, 260, 261, 262, 269, 270],
        );
        println!("  ✓ Executive Summary show: 9 slides");

        pres.create_custom_show(
            "Financial Review",
            vec![256, 257, 258, 259, 267, 268, 269, 270],
        );
        println!("  ✓ Financial Review show: 8 slides");

        pres.create_custom_show("Technical Review", vec![256, 262, 263, 264, 265, 270]);
        println!("  ✓ Technical Review show: 6 slides");

        pres.create_custom_show("All Hands Meeting", (256..=270).collect::<Vec<_>>());
        println!("  ✓ All Hands Meeting show: 15 slides");

        println!("  ✓ Generated {} custom shows", pres.custom_shows().len());
    }
    pkg.save("customshow_executive.pptx")?;
    println!("  ✓ Saved: customshow_executive.pptx\n");

    Ok(())
}

/// Test 2: Training Presentation
/// Custom shows for different roles and skill levels
fn test_training_presentation() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 2: Creating training presentation with role-specific shows...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Create 20 training slides
        let training_topics = vec![
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

        for (_i, title) in training_topics.iter().enumerate() {
            let slide = pres.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(
                &format!("Training content for {}", title),
                914400,
                1828800,
                7315200,
                3000000,
            );
        }

        // Create role-specific custom shows
        pres.create_custom_show("New Employee Orientation", vec![256, 257, 258, 275]);
        println!("  ✓ New Employee Orientation: 4 slides");

        pres.create_custom_show("Manager Training", vec![256, 257, 258, 259, 260, 261, 275]);
        println!("  ✓ Manager Training: 7 slides");

        pres.create_custom_show("HR Department", vec![256, 257, 258, 262, 263, 264, 275]);
        println!("  ✓ HR Department: 7 slides");

        pres.create_custom_show("IT Department", vec![256, 257, 258, 265, 266, 267, 275]);
        println!("  ✓ IT Department: 7 slides");

        pres.create_custom_show("Sales Team", vec![256, 257, 268, 269, 270, 275]);
        println!("  ✓ Sales Team: 6 slides");

        pres.create_custom_show("Support Team", vec![256, 257, 271, 272, 273, 275]);
        println!("  ✓ Support Team: 6 slides");

        println!("  ✓ Generated {} custom shows", pres.custom_shows().len());
    }
    pkg.save("customshow_training.pptx")?;
    println!("  ✓ Saved: customshow_training.pptx\n");

    Ok(())
}

/// Test 3: Product Demo
/// Feature-specific shows for different customer interests
fn test_product_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 3: Creating product demo with feature-specific shows...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Create product demo slides
        let demo_slides = vec![
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

        for (_i, title) in demo_slides.iter().enumerate() {
            let slide = pres.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(
                &format!("Demo content: {}", title),
                914400,
                1828800,
                7315200,
                3500000,
            );
        }

        // Create interest-based custom shows
        pres.create_custom_show("Quick Overview", vec![256, 257, 267, 271]);
        println!("  ✓ Quick Overview: 4 slides (5 minutes)");

        pres.create_custom_show(
            "Analytics Focus",
            vec![256, 257, 258, 259, 260, 267, 269, 270, 271],
        );
        println!("  ✓ Analytics Focus: 9 slides");

        pres.create_custom_show(
            "Technical Deep Dive",
            vec![256, 261, 262, 263, 266, 267, 271],
        );
        println!("  ✓ Technical Deep Dive: 7 slides");

        pres.create_custom_show("Business Value", vec![256, 257, 267, 269, 270, 271]);
        println!("  ✓ Business Value: 6 slides");

        pres.create_custom_show("Mobile-First", vec![256, 263, 264, 267, 271]);
        println!("  ✓ Mobile-First: 5 slides");

        pres.create_custom_show("Full Demo", (256..=271).collect::<Vec<_>>());
        println!("  ✓ Full Demo: 16 slides (complete)");

        println!("  ✓ Generated {} custom shows", pres.custom_shows().len());
    }
    pkg.save("customshow_product_demo.pptx")?;
    println!("  ✓ Saved: customshow_product_demo.pptx\n");

    Ok(())
}

/// Test 4: Conference Presentation
/// Time-based shows for different session lengths
fn test_conference_presentation() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 4: Creating conference presentation with time-based shows...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Create conference presentation slides
        let conference_content = vec![
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

        for (_i, title) in conference_content.iter().enumerate() {
            let slide = pres.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(
                &format!("Conference content: {}", title),
                914400,
                1828800,
                7315200,
                3200000,
            );
        }

        // Create time-based custom shows
        pres.create_custom_show("Lightning Talk (5 min)", vec![256, 257, 259, 274]);
        println!("  ✓ Lightning Talk: 4 slides (5 minutes)");

        pres.create_custom_show(
            "Short Session (15 min)",
            vec![256, 257, 258, 259, 260, 264, 267, 274],
        );
        println!("  ✓ Short Session: 8 slides (15 minutes)");

        pres.create_custom_show(
            "Standard Talk (30 min)",
            vec![
                256, 257, 258, 259, 260, 261, 264, 265, 266, 267, 269, 271, 274,
            ],
        );
        println!("  ✓ Standard Talk: 13 slides (30 minutes)");

        pres.create_custom_show("Extended Session (45 min)", (256..=272).collect::<Vec<_>>());
        println!("  ✓ Extended Session: 17 slides (45 minutes)");

        pres.create_custom_show("Workshop (90 min)", (256..=274).collect::<Vec<_>>());
        println!("  ✓ Workshop: 19 slides (90 minutes)");

        println!(
            "  ✓ Generated {} time-based shows",
            pres.custom_shows().len()
        );
    }
    pkg.save("customshow_conference.pptx")?;
    println!("  ✓ Saved: customshow_conference.pptx\n");

    Ok(())
}
