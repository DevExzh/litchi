//! Comprehensive Handout Master Tests
//!
//! This example generates multiple PPTX files to test all handout master features:
//! - Different layouts (1, 2, 3, 4, 6, 9 slides per page)
//! - Header/footer visibility options
//! - Background colors
//! - Combined with other presentation features

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::handout::{HandoutLayout, HandoutMaster};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive Handout Master Tests ===\n");

    // Test 1: Different Handout Layouts
    test_handout_layouts()?;

    // Test 2: Background Colors
    test_background_colors()?;

    // Test 3: Multi-slide presentation with handout
    test_multi_slide_handout()?;

    // Test 4: Full-featured presentation with handout
    test_full_featured_handout()?;

    println!("\n=== All handout tests complete! ===");
    println!("\nTo verify in PowerPoint:");
    println!("  1. Open each file");
    println!("  2. Go to File > Print");
    println!("  3. Select 'Handouts' in Print Layout");
    println!("  4. Check the layout matches the expected slides per page");

    Ok(())
}

/// Test 1: Create files with each handout layout type
fn test_handout_layouts() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 1: Testing different handout layouts...");

    let layouts = [
        (HandoutLayout::OneSlide, "1_slide", "1 slide per page"),
        (HandoutLayout::TwoSlides, "2_slides", "2 slides per page"),
        (
            HandoutLayout::ThreeSlides,
            "3_slides",
            "3 slides per page (with lines)",
        ),
        (HandoutLayout::FourSlides, "4_slides", "4 slides per page"),
        (HandoutLayout::SixSlides, "6_slides", "6 slides per page"),
        (HandoutLayout::NineSlides, "9_slides", "9 slides per page"),
    ];

    for (layout, suffix, description) in layouts {
        let mut pkg = Package::new()?;
        {
            let pres = pkg.presentation_mut()?;

            // Add multiple slides to make handout meaningful
            for i in 1..=6 {
                let slide = pres.add_slide()?;
                slide.set_title(&format!("Slide {} - {}", i, description));
                slide.add_text_box(
                    &format!("Content for slide {}\nTesting layout: {}", i, description),
                    914400,
                    1828800,
                    7315200,
                    1500000,
                );
            }

            // Set handout master with specific layout
            let mut handout = HandoutMaster::new();
            handout.layout = layout;
            pres.set_handout_master(handout);
        }
        let filename = format!("handout_layout_{}.pptx", suffix);
        pkg.save(&filename)?;
        println!("  ✓ Saved: {} ({})", filename, description);
    }

    Ok(())
}

/// Test 2: Test background colors
fn test_background_colors() -> Result<(), Box<dyn std::error::Error>> {
    println!("\nTest 2: Testing background colors...");

    let colors = [
        ("FFFFFF", "white"),
        ("F0F0F0", "light_gray"),
        ("E6F3FF", "light_blue"),
        ("FFF0E6", "light_orange"),
        ("E6FFE6", "light_green"),
        ("FFE6F0", "light_pink"),
    ];

    for (color_hex, color_name) in colors {
        let mut pkg = Package::new()?;
        {
            let pres = pkg.presentation_mut()?;

            for i in 1..=4 {
                let slide = pres.add_slide()?;
                slide.set_title(&format!("Slide {} - {} background", i, color_name));
                slide.add_text_box(
                    &format!("Testing {} (#{}) handout background", color_name, color_hex),
                    914400,
                    1828800,
                    7315200,
                    914400,
                );
            }

            let mut handout = HandoutMaster::new();
            handout.layout = HandoutLayout::FourSlides;
            handout.background_color = Some(color_hex.to_string());
            pres.set_handout_master(handout);
        }
        let filename = format!("handout_bg_{}.pptx", color_name);
        pkg.save(&filename)?;
        println!("  ✓ Saved: {} (#{} background)", filename, color_hex);
    }

    Ok(())
}

/// Test 3: Multi-slide presentation with different content types
fn test_multi_slide_handout() -> Result<(), Box<dyn std::error::Error>> {
    println!("\nTest 3: Testing multi-slide presentation with handout...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Create 12 slides with varied content
        for i in 1..=12 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Chapter {} Content", i));

            // Add a text box with content
            slide.add_text_box(
                &format!(
                    "This is slide {} with sample text content.\n\n\
                     Handouts are useful for:\n\
                     • Taking notes during presentations\n\
                     • Reference material for attendees\n\
                     • Print-friendly version of slides\n\n\
                     Page {} of 12",
                    i, i
                ),
                500_000,
                1_800_000,
                8_000_000,
                4_000_000,
            );
        }

        // Set handout with 6 slides per page (good for 12 slides = 2 pages)
        let mut handout = HandoutMaster::new();
        handout.layout = HandoutLayout::SixSlides;
        pres.set_handout_master(handout);
    }
    pkg.save("handout_12_slides.pptx")?;
    println!("  ✓ Saved: handout_12_slides.pptx (12 slides, 6 per handout page)");

    Ok(())
}

/// Test 4: Full-featured presentation with handout
fn test_full_featured_handout() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::ooxml::pptx::animations::AnimationEffect;

    println!("\nTest 4: Testing full-featured presentation with handout...");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Slide 1: Title slide
        {
            let slide = pres.add_slide()?;
            slide.set_title("Welcome to the Presentation");
            slide.add_text_box(
                "A comprehensive test of handout features\n\nPresented by: Test Author",
                914400,
                2500000,
                7315200,
                1500000,
            );
        }

        // Slide 2: Agenda
        {
            let slide = pres.add_slide()?;
            slide.set_title("Agenda");
            slide.add_text_box(
                "• Introduction\n• Main Content\n• Demonstrations\n• Conclusion",
                914400,
                1828800,
                7315200,
                2500000,
            );
            slide.add_animation(3, AnimationEffect::FlyIn);
        }

        // Slide 3: Key Points
        {
            let slide = pres.add_slide()?;
            slide.set_title("Key Points");
            slide.add_text_box(
                "Important information goes here.\n\n\
                 This slide demonstrates:\n\
                 - Text formatting\n\
                 - Bullet points\n\
                 - Multiple paragraphs",
                914400,
                1828800,
                7315200,
                3000000,
            );
        }

        // Slide 4: Data Overview
        {
            let slide = pres.add_slide()?;
            slide.set_title("Data Overview");
            slide.add_rectangle(
                1500000,
                2000000,
                6000000,
                3000000,
                Some("4472C4".to_string()),
            ); // Blue fill
        }

        // Slide 5: Live Demo
        {
            let slide = pres.add_slide()?;
            slide.set_title("Live Demo");
            slide.add_text_box(
                "Interactive demonstration section.\n\n\
                 Follow along with the handouts to take notes.",
                914400,
                1828800,
                7315200,
                2500000,
            );
            slide.add_animation(3, AnimationEffect::Fade);
        }

        // Slide 6: Step-by-Step Guide
        {
            let slide = pres.add_slide()?;
            slide.set_title("Step-by-Step Guide");
            slide.add_text_box(
                "Step 1: Open the application\n\
                 Step 2: Navigate to settings\n\
                 Step 3: Configure options\n\
                 Step 4: Save and apply",
                914400,
                1828800,
                7315200,
                3000000,
            );
        }

        // Slide 7: Summary
        {
            let slide = pres.add_slide()?;
            slide.set_title("Summary");
            slide.add_text_box(
                "Key takeaways from today's presentation:\n\n\
                 ✓ Feature 1 explained\n\
                 ✓ Feature 2 demonstrated\n\
                 ✓ Best practices covered",
                914400,
                1828800,
                7315200,
                2500000,
            );
        }

        // Slide 8: Thank You
        {
            let slide = pres.add_slide()?;
            slide.set_title("Thank You!");
            slide.add_text_box(
                "Questions & Discussion\n\n\
                 Contact: example@email.com\n\
                 Website: www.example.com",
                914400,
                2000000,
                7315200,
                2500000,
            );
        }

        // Add sections
        pres.add_section("Introduction", vec![256, 257]);
        pres.add_section("Main Content", vec![258, 259]);
        pres.add_section("Demonstrations", vec![260, 261]);
        pres.add_section("Conclusion", vec![262, 263]);

        // Configure handout master with all features including header/footer
        let handout = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides) // 3 slides with lines for notes
            .with_background_color("F5F5F5") // Light gray background
            .with_header("Confidential - Internal Use Only")
            .with_footer("© 2024 Example Corp")
            .with_slide_numbers()
            .with_date_time();
        pres.set_handout_master(handout);
    }
    pkg.save("handout_full_featured.pptx")?;
    println!(
        "  ✓ Saved: handout_full_featured.pptx (8 slides with sections, animations, handout with header/footer)"
    );

    // Also create a header/footer only test
    let mut pkg2 = Package::new()?;
    {
        let pres = pkg2.presentation_mut()?;
        for i in 1..=4 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));
        }

        let handout = HandoutMaster::new()
            .with_header("Header Text Here")
            .with_footer("Footer Text Here")
            .with_slide_numbers()
            .with_fixed_date("December 5, 2024");
        pres.set_handout_master(handout);
    }
    pkg2.save("handout_header_footer_test.pptx")?;
    println!(
        "  ✓ Saved: handout_header_footer_test.pptx (with header, footer, slide numbers, fixed date)"
    );

    // Test auto-date only (to isolate issue)
    let mut pkg3 = Package::new()?;
    {
        let pres = pkg3.presentation_mut()?;
        for i in 1..=4 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));
        }

        // Same as full_featured but NO sections/animations
        let handout = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides)
            .with_background_color("F5F5F5")
            .with_header("Confidential - Internal Use Only")
            .with_footer("© 2024 Example Corp")
            .with_slide_numbers()
            .with_date_time();
        pres.set_handout_master(handout);
    }
    pkg3.save("handout_auto_date_test.pptx")?;
    println!(
        "  ✓ Saved: handout_auto_date_test.pptx (same handout as full_featured, no sections/animations)"
    );

    // BINARY SEARCH TESTS - to isolate full_featured issue
    println!("\n=== Binary Search Tests ===");

    // Test 1: 8 slides + handout (NO sections, NO animations)
    test_8slides_handout()?;

    // Test 2: 8 slides + sections + handout (NO animations)
    test_8slides_sections_handout()?;

    // Test 3: 8 slides + animations + handout (NO sections)
    test_8slides_animations_handout()?;

    // Test 4: 8 slides + sections + animations (NO handout)
    test_8slides_sections_animations_no_handout()?;

    Ok(())
}

/// Test 1: 8 slides + handout (NO sections, NO animations)
fn test_8slides_handout() -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        for i in 1..=8 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));
        }

        let handout = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides)
            .with_background_color("F5F5F5")
            .with_header("Confidential - Internal Use Only")
            .with_footer("© 2024 Example Corp")
            .with_slide_numbers()
            .with_date_time();
        pres.set_handout_master(handout);
    }
    pkg.save("test_8slides_handout.pptx")?;
    println!("  ✓ Saved: test_8slides_handout.pptx (8 slides + handout)");
    Ok(())
}

/// Test 2: 8 slides + sections + handout (NO animations)
fn test_8slides_sections_handout() -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        for i in 1..=8 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));
        }

        // Add sections (same as full_featured)
        pres.add_section("Introduction", vec![256, 257]);
        pres.add_section("Main Content", vec![258, 259]);
        pres.add_section("Demonstrations", vec![260, 261]);
        pres.add_section("Conclusion", vec![262, 263]);

        let handout = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides)
            .with_background_color("F5F5F5")
            .with_header("Confidential - Internal Use Only")
            .with_footer("© 2024 Example Corp")
            .with_slide_numbers()
            .with_date_time();
        pres.set_handout_master(handout);
    }
    pkg.save("test_8slides_sections_handout.pptx")?;
    println!("  ✓ Saved: test_8slides_sections_handout.pptx (8 slides + sections + handout)");
    Ok(())
}

/// Test 3: 8 slides + animations + handout (NO sections)
fn test_8slides_animations_handout() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::ooxml::pptx::animations::AnimationEffect;

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        for i in 1..=8 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));

            // Add text box so shape ID 3 exists
            slide.add_text_box(
                &format!("Content for slide {}", i),
                914400,
                1828800,
                7315200,
                2500000,
            );

            // Add animation to slides 2 and 5 (same as full_featured)
            if i == 2 || i == 5 {
                slide.add_animation(3, AnimationEffect::FlyIn);
            }
        }

        let handout = HandoutMaster::new()
            .with_layout(HandoutLayout::ThreeSlides)
            .with_background_color("F5F5F5")
            .with_header("Confidential - Internal Use Only")
            .with_footer("© 2024 Example Corp")
            .with_slide_numbers()
            .with_date_time();
        pres.set_handout_master(handout);
    }
    pkg.save("test_8slides_animations_handout.pptx")?;
    println!("  ✓ Saved: test_8slides_animations_handout.pptx (8 slides + animations + handout)");
    Ok(())
}

/// Test 4: 8 slides + sections + animations (NO handout)
fn test_8slides_sections_animations_no_handout() -> Result<(), Box<dyn std::error::Error>> {
    use litchi::ooxml::pptx::animations::AnimationEffect;

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        for i in 1..=8 {
            let slide = pres.add_slide()?;
            slide.set_title(&format!("Slide {}", i));

            // Add text box so shape ID 3 exists
            slide.add_text_box(
                &format!("Content for slide {}", i),
                914400,
                1828800,
                7315200,
                2500000,
            );

            if i == 2 || i == 5 {
                slide.add_animation(3, AnimationEffect::FlyIn);
            }
        }

        pres.add_section("Introduction", vec![256, 257]);
        pres.add_section("Main Content", vec![258, 259]);
        pres.add_section("Demonstrations", vec![260, 261]);
        pres.add_section("Conclusion", vec![262, 263]);

        // NO handout master
    }
    pkg.save("test_8slides_sections_animations_no_handout.pptx")?;
    println!(
        "  ✓ Saved: test_8slides_sections_animations_no_handout.pptx (8 slides + sections + animations, NO handout)"
    );
    Ok(())
}
