//! Isolates the PPTX corruption issue by testing combinations

use litchi::ooxml::pptx::animations::AnimationEffect;
use litchi::ooxml::pptx::handout::{HandoutLayout, HandoutMaster};
use litchi::ooxml::pptx::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    // Test A: Animations + Sections (no custom shows, no handout)
    test_anim_sections()?;

    // Test B: Animations + Custom Shows (no sections, no handout)
    test_anim_custom_shows()?;

    // Test C: Sections + Custom Shows (no animations, no handout)
    test_sections_custom_shows()?;

    // Test D: Sections + Custom Shows + Handout (no animations)
    test_sections_custom_shows_handout()?;

    // Test E: Animations only (no sections, no custom shows, no handout)
    test_animations_only()?;

    println!("\nTests complete! Check which files fail in PowerPoint.");

    Ok(())
}

fn test_anim_sections() -> Result<(), Box<dyn Error>> {
    println!("Test A: Animations + Sections...");
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
    }
    pkg.save("test_A_anim_sections.pptx")?;
    println!("  Saved: test_A_anim_sections.pptx");
    Ok(())
}

fn test_anim_custom_shows() -> Result<(), Box<dyn Error>> {
    println!("Test B: Animations + Custom Shows...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");
        slide1.add_text_box("Text", 914400, 1828800, 7315200, 914400);
        slide1.add_animation(3, AnimationEffect::Fade);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.create_custom_show("My Show", vec![256, 257]);
    }
    pkg.save("test_B_anim_custom_shows.pptx")?;
    println!("  Saved: test_B_anim_custom_shows.pptx");
    Ok(())
}

fn test_sections_custom_shows() -> Result<(), Box<dyn Error>> {
    println!("Test C: Sections + Custom Shows...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.add_section("Section 1", vec![256]);
        pres.add_section("Section 2", vec![257]);

        pres.create_custom_show("My Show", vec![256, 257]);
    }
    pkg.save("test_C_sections_custom_shows.pptx")?;
    println!("  Saved: test_C_sections_custom_shows.pptx");
    Ok(())
}

fn test_sections_custom_shows_handout() -> Result<(), Box<dyn Error>> {
    println!("Test D: Sections + Custom Shows + Handout...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1");

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2");

        pres.add_section("Section 1", vec![256]);
        pres.add_section("Section 2", vec![257]);

        pres.create_custom_show("My Show", vec![256, 257]);

        let mut handout = HandoutMaster::new();
        handout.layout = HandoutLayout::SixSlides;
        pres.set_handout_master(handout);
    }
    pkg.save("test_D_sections_custom_shows_handout.pptx")?;
    println!("  Saved: test_D_sections_custom_shows_handout.pptx");
    Ok(())
}

fn test_animations_only() -> Result<(), Box<dyn Error>> {
    println!("Test E: Animations only (multiple)...");
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Animation Test");
        slide1.add_text_box("Text 1", 914400, 1828800, 3657600, 914400);
        slide1.add_text_box("Text 2", 914400, 2971800, 3657600, 914400);
        slide1.add_text_box("Text 3", 914400, 4114800, 3657600, 914400);

        // Add multiple animations
        slide1.add_animation(3, AnimationEffect::Fade);
        slide1.add_animation(4, AnimationEffect::FlyIn);
        slide1.add_animation(5, AnimationEffect::Wipe);
    }
    pkg.save("test_E_animations_only.pptx")?;
    println!("  Saved: test_E_animations_only.pptx");
    Ok(())
}
