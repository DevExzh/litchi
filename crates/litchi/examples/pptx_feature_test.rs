//! PPTX Feature Isolation Test
//!
//! Creates multiple PPTX files to test each feature individually.

use litchi::ooxml::pptx::animations::AnimationEffect;
use litchi::ooxml::pptx::handout::{HandoutLayout, HandoutMaster};
use litchi::ooxml::pptx::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    // Test 1: Basic slides only
    test_basic()?;

    // Test 2: With animations
    test_animations()?;

    // Test 3: With sections
    test_sections()?;

    // Test 4: With custom shows
    test_custom_shows()?;

    // Test 5: With handout master
    test_handout()?;

    // Test 6: All features
    test_all_features()?;

    println!("\nAll tests complete! Open each file in PowerPoint to find which one fails.");

    Ok(())
}

fn test_basic() -> Result<(), Box<dyn Error>> {
    println!("Test 1: Basic slides...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Basic Test");
        slide.add_text_box("Hello World", 914400, 1828800, 7315200, 914400);
    }
    pkg.save("test_1_basic.pptx")?;
    println!("  Saved: test_1_basic.pptx");
    Ok(())
}

fn test_animations() -> Result<(), Box<dyn Error>> {
    println!("Test 2: With animations...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Animation Test");
        slide.add_text_box("Animated text", 914400, 1828800, 7315200, 914400);
        slide.add_animation(3, AnimationEffect::Fade);
    }
    pkg.save("test_2_animations.pptx")?;
    println!("  Saved: test_2_animations.pptx");
    Ok(())
}

fn test_sections() -> Result<(), Box<dyn Error>> {
    println!("Test 3: With sections...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");
        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.add_section("Section 1", vec![256]);
        pres.add_section("Section 2", vec![257]);
    }
    pkg.save("test_3_sections.pptx")?;
    println!("  Saved: test_3_sections.pptx");
    Ok(())
}

fn test_custom_shows() -> Result<(), Box<dyn Error>> {
    println!("Test 4: With custom shows...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");
        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.create_custom_show("My Show", vec![256, 257]);
    }
    pkg.save("test_4_custom_shows.pptx")?;
    println!("  Saved: test_4_custom_shows.pptx");
    Ok(())
}

fn test_handout() -> Result<(), Box<dyn Error>> {
    println!("Test 5: With handout master...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Handout Test");

        let mut handout = HandoutMaster::new();
        handout.layout = HandoutLayout::SixSlides;
        pres.set_handout_master(handout);
    }
    pkg.save("test_5_handout.pptx")?;
    println!("  Saved: test_5_handout.pptx");
    Ok(())
}

fn test_all_features() -> Result<(), Box<dyn Error>> {
    println!("Test 6: All features...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");
        slide1.add_text_box("Text", 914400, 1828800, 7315200, 914400);
        slide1.add_animation(3, AnimationEffect::Fade);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.add_section("Section 1", vec![256]);
        pres.add_section("Section 2", vec![257]);

        pres.create_custom_show("My Show", vec![256, 257]);

        let mut handout = HandoutMaster::new();
        handout.layout = HandoutLayout::SixSlides;
        pres.set_handout_master(handout);
    }
    pkg.save("test_6_all_features.pptx")?;
    println!("  Saved: test_6_all_features.pptx");
    Ok(())
}
