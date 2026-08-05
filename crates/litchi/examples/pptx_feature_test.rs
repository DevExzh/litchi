//! PPTX feature isolation test.
//!
//! Each output file contains a small inspectable summary deck. Feature-specific
//! PresentationML is built and round-tripped through its typed owner before it
//! is summarized in the package, keeping the package facade focused on slides.

use litchi_pptx::animations::{Effect, EffectInstance, Sequence, Trigger};
use litchi_pptx::presentation_properties::metadata::custom_show::{List as ShowList, Show};
use litchi_pptx::presentation_properties::metadata::handout::{Layout, Master};
use litchi_pptx::presentation_properties::metadata::sections::{List as SectionList, Section};
use litchi_pptx::{MutablePresentation, Package};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_600_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 3_800_000;

fn main() -> Result<(), Box<dyn StdError>> {
    test_basic()?;
    test_animations()?;
    test_sections()?;
    test_custom_shows()?;
    test_handout()?;
    test_all_features()?;

    println!("\nAll tests complete! Open each file in PowerPoint to inspect the summaries.");
    Ok(())
}

fn test_basic() -> Result<(), Box<dyn StdError>> {
    println!("Test 1: Basic slides...");
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "Basic Test",
            "Basic slide and text-box package authoring.",
        )?;
    }
    save_and_reopen(package, "test_1_basic.pptx")?;
    Ok(())
}

fn test_animations() -> Result<(), Box<dyn StdError>> {
    println!("Test 2: With animations...");
    let animation_xml = animation_demo()?;
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "Animation Test",
            &format!(
                "Typed animation sequence\nCanonical XML: {} bytes",
                animation_xml.len()
            ),
        )?;
    }
    save_and_reopen(package, "test_2_animations.pptx")?;
    Ok(())
}

fn test_sections() -> Result<(), Box<dyn StdError>> {
    println!("Test 3: With sections...");
    let sections_xml = sections_demo()?;
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "Sections Test",
            &format!(
                "Typed section list\nCanonical XML: {} bytes",
                sections_xml.len()
            ),
        )?;
    }
    save_and_reopen(package, "test_3_sections.pptx")?;
    Ok(())
}

fn test_custom_shows() -> Result<(), Box<dyn StdError>> {
    println!("Test 4: With custom shows...");
    let shows_xml = custom_shows_demo();
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "Custom Shows Test",
            &format!(
                "Typed custom-show list\nCanonical XML: {} bytes",
                shows_xml.len()
            ),
        )?;
    }
    save_and_reopen(package, "test_4_custom_shows.pptx")?;
    Ok(())
}

fn test_handout() -> Result<(), Box<dyn StdError>> {
    println!("Test 5: With handout master...");
    let handout_xml = handout_demo();
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "Handout Test",
            &format!(
                "Typed handout master\nCanonical XML: {} bytes",
                handout_xml.len()
            ),
        )?;
    }
    save_and_reopen(package, "test_5_handout.pptx")?;
    Ok(())
}

fn test_all_features() -> Result<(), Box<dyn StdError>> {
    println!("Test 6: All features...");
    let animation_xml = animation_demo()?;
    let sections_xml = sections_demo()?;
    let shows_xml = custom_shows_demo();
    let handout_xml = handout_demo();

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        add_summary_slide(
            presentation,
            "All Features",
            &format!(
                "Typed PresentationML feature owners\n\nAnimations: {} bytes\nSections: {} bytes\nCustom shows: {} bytes\nHandout: {} bytes",
                animation_xml.len(),
                sections_xml.len(),
                shows_xml.len(),
                handout_xml.len(),
            ),
        )?;
    }
    save_and_reopen(package, "test_6_all_features.pptx")?;
    Ok(())
}

fn add_summary_slide(
    presentation: &mut MutablePresentation,
    title: &str,
    body: &str,
) -> Result<(), Box<dyn StdError>> {
    let slide = presentation.add_slide()?;
    slide.set_title(title);
    slide.add_text_box(body, X, Y, WIDTH, HEIGHT);
    Ok(())
}

fn save_and_reopen(mut package: Package, path: &str) -> Result<(), Box<dyn StdError>> {
    package.save(path)?;
    let reopened = Package::open(path)?;
    assert_eq!(reopened.presentation()?.slide_count()?, 1);
    println!("  Saved and reopened: {path}");
    Ok(())
}

fn animation_demo() -> Result<String, Box<dyn StdError>> {
    let mut sequence = Sequence::new();
    sequence.add(
        EffectInstance::new(3, Effect::Fade)
            .with_trigger(Trigger::OnClick)
            .with_duration_ms(500),
    );
    sequence.add(
        EffectInstance::new(4, Effect::FlyIn)
            .with_trigger(Trigger::AfterPrevious)
            .with_duration_ms(750),
    );
    let xml = sequence.to_xml();
    let parsed = Sequence::parse_timing_xml(&xml)?;
    assert_eq!(parsed.animations.len(), 2);
    Ok(xml)
}

fn sections_demo() -> Result<String, Box<dyn StdError>> {
    let mut sections = SectionList::new();
    sections.add_section(Section::new("Section 1", "section-1").with_slides([256]));
    sections.add_section(Section::new("Section 2", "section-2").with_slides([257]));
    Ok(sections.to_xml()?)
}

fn custom_shows_demo() -> String {
    let mut shows = ShowList::new();
    shows.add(Show::new(1, "My Show").with_slides(vec![256, 257]));
    shows.to_xml()
}

fn handout_demo() -> String {
    Master::new()
        .with_layout(Layout::SixSlides)
        .with_header("Feature Test")
        .with_footer("Created with Litchi")
        .with_slide_numbers()
        .to_xml()
}
