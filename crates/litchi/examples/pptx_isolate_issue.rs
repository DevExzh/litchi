//! Isolate PresentationML combinations at the typed-owner and package boundaries.
//!
//! Each case keeps the original animation/section/custom-show/handout
//! combination, but demonstrates those features through their current typed
//! XML owners. The accompanying package contains a summary slide deck so the
//! cases remain easy to inspect independently.

use litchi_pptx::animations::{Effect, EffectInstance, Sequence, Trigger};
use litchi_pptx::presentation_properties::metadata::custom_show::{List as ShowList, Show};
use litchi_pptx::presentation_properties::metadata::handout::{Layout, Master};
use litchi_pptx::presentation_properties::metadata::sections::{List as SectionList, Section};
use litchi_pptx::{MutablePresentation, Package};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_828_800;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 914_400;

fn main() -> Result<(), Box<dyn StdError>> {
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

fn test_anim_sections() -> Result<(), Box<dyn StdError>> {
    println!("Test A: Animations + Sections...");
    let animations = animations_xml(&[(3, Effect::Fade)])?;
    let sections = sections_xml()?;
    write_case(
        "test_A_anim_sections.pptx",
        "Animations + Sections",
        &format!(
            "Animations XML: {} bytes\nSections XML: {} bytes",
            animations.len(),
            sections.len()
        ),
        2,
    )?;
    Ok(())
}

fn test_anim_custom_shows() -> Result<(), Box<dyn StdError>> {
    println!("Test B: Animations + Custom Shows...");
    let animations = animations_xml(&[(3, Effect::Fade)])?;
    let shows = custom_shows_xml();
    write_case(
        "test_B_anim_custom_shows.pptx",
        "Animations + Custom Shows",
        &format!(
            "Animations XML: {} bytes\nCustom shows XML: {} bytes",
            animations.len(),
            shows.len()
        ),
        2,
    )?;
    Ok(())
}

fn test_sections_custom_shows() -> Result<(), Box<dyn StdError>> {
    println!("Test C: Sections + Custom Shows...");
    let sections = sections_xml()?;
    let shows = custom_shows_xml();
    write_case(
        "test_C_sections_custom_shows.pptx",
        "Sections + Custom Shows",
        &format!(
            "Sections XML: {} bytes\nCustom shows XML: {} bytes",
            sections.len(),
            shows.len()
        ),
        2,
    )?;
    Ok(())
}

fn test_sections_custom_shows_handout() -> Result<(), Box<dyn StdError>> {
    println!("Test D: Sections + Custom Shows + Handout...");
    let sections = sections_xml()?;
    let shows = custom_shows_xml();
    let handout = handout_xml();
    write_case(
        "test_D_sections_custom_shows_handout.pptx",
        "Sections + Custom Shows + Handout",
        &format!(
            "Sections XML: {} bytes\nCustom shows XML: {} bytes\nHandout XML: {} bytes",
            sections.len(),
            shows.len(),
            handout.len()
        ),
        2,
    )?;
    Ok(())
}

fn test_animations_only() -> Result<(), Box<dyn StdError>> {
    println!("Test E: Animations only (multiple)...");
    let animations = animations_xml(&[(3, Effect::Fade), (4, Effect::FlyIn), (5, Effect::Wipe)])?;
    write_case(
        "test_E_animations_only.pptx",
        "Animations only",
        &format!("Animations XML: {} bytes", animations.len()),
        1,
    )?;
    Ok(())
}

fn animations_xml(effects: &[(u32, Effect)]) -> Result<String, Box<dyn StdError>> {
    let mut sequence = Sequence::new();
    for (shape_id, effect) in effects {
        sequence.add(
            EffectInstance::new(*shape_id, effect.clone())
                .with_trigger(Trigger::OnClick)
                .with_duration_ms(500),
        );
    }

    let xml = sequence.to_xml();
    let parsed = Sequence::parse_timing_xml(&xml)?;
    assert_eq!(parsed.animations.len(), effects.len());
    Ok(xml)
}

fn sections_xml() -> Result<String, Box<dyn StdError>> {
    let mut sections = SectionList::new();
    sections.add_section(Section::new("Section 1", "section-1").with_slides([256]));
    sections.add_section(Section::new("Section 2", "section-2").with_slides([257]));
    Ok(sections.to_xml()?)
}

fn custom_shows_xml() -> String {
    let mut shows = ShowList::new();
    shows.add(Show::new(1, "My Show").with_slides(vec![256, 257]));
    let xml = shows.to_xml();
    let parsed = ShowList::parse_xml(&xml).expect("valid typed custom-show model");
    assert_eq!(parsed.len(), 1);
    xml
}

fn handout_xml() -> String {
    let handout = Master::new()
        .with_layout(Layout::SixSlides)
        .with_header("Issue isolation")
        .with_footer("Litchi");
    assert_eq!(handout.layout, Layout::SixSlides);
    let xml = handout.to_xml();
    assert!(xml.contains("handoutMaster"));
    xml
}

fn write_case(
    path: &str,
    title: &str,
    body: &str,
    slide_count: usize,
) -> Result<(), Box<dyn StdError>> {
    let mut package = Package::new()?;
    {
        let presentation: &mut MutablePresentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();
        for index in 1..=slide_count {
            let slide = presentation.add_slide()?;
            let slide_title = if index == 1 {
                title.to_owned()
            } else {
                format!("{title} — slide {index}")
            };
            slide.set_title(&slide_title);
            slide.add_text_box(body, X, Y, WIDTH, HEIGHT);
        }
    }
    package.save(path)?;
    let reopened = Package::open(path)?;
    assert_eq!(reopened.presentation()?.slide_count()?, slide_count);
    println!("  Saved: {path}");
    Ok(())
}
