//! Comprehensive PresentationML handout, animation, and section showcase.
//!
//! The standalone `litchi-pptx` facade exposes these features as typed owners
//! with bounded XML codecs.  Handout masters, timing trees, and sections are
//! demonstrated at their model/XML boundary, while the package writer creates
//! and reopens a summary deck for every case.
//!
//! Run with: `cargo run --example pptx_handout_comprehensive --features ooxml`

use litchi_pptx::Package;
use litchi_pptx::animations::{Effect, EffectInstance, Sequence, Trigger};
use litchi_pptx::presentation_properties::metadata::handout::{Layout, Master};
use litchi_pptx::presentation_properties::metadata::sections::{List as SectionList, Section};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_800_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 3_600_000;

fn main() -> Result<(), Box<dyn StdError>> {
    println!("=== Comprehensive PresentationML feature showcase ===\n");

    test_handout_layouts()?;
    test_background_colors()?;
    test_multi_slide_handout()?;
    test_full_featured_handout()?;

    println!("\n=== All handout, animation, and section cases complete! ===");
    Ok(())
}

/// Test 1: exercise every supported print layout.
fn test_handout_layouts() -> Result<(), Box<dyn StdError>> {
    println!("Test 1: typed handout layouts...");

    let layouts = [
        (Layout::OneSlide, "1_slide", "1 slide per page"),
        (Layout::TwoSlides, "2_slides", "2 slides per page"),
        (
            Layout::ThreeSlides,
            "3_slides",
            "3 slides per page with notes lines",
        ),
        (Layout::FourSlides, "4_slides", "4 slides per page"),
        (Layout::SixSlides, "6_slides", "6 slides per page"),
        (Layout::NineSlides, "9_slides", "9 slides per page"),
    ];

    for (layout, suffix, description) in layouts {
        let handout = Master::new().with_layout(layout);
        let handout_xml = handout.to_xml();
        assert!(handout_xml.contains(layout.as_str()));
        write_summary_package(
            &format!("handout_layout_{suffix}.pptx"),
            &format!("Handout layout: {description}"),
            &format!("{}\nXML bytes: {}", layout.print_what(), handout_xml.len()),
            6,
        )?;
        println!("  ✓ {description}");
    }

    Ok(())
}

/// Test 2: preserve background and header/footer configuration in typed XML.
fn test_background_colors() -> Result<(), Box<dyn StdError>> {
    println!("\nTest 2: typed handout background and visibility settings...");

    let colors = [
        ("FFFFFF", "white"),
        ("F0F0F0", "light_gray"),
        ("E6F3FF", "light_blue"),
        ("FFF0E6", "light_orange"),
        ("E6FFE6", "light_green"),
        ("FFE6F0", "light_pink"),
    ];

    for (color, name) in colors {
        let handout = Master::new()
            .with_layout(Layout::FourSlides)
            .with_background_color(color)
            .with_header(format!("{name} header"))
            .with_footer("Printed handout")
            .with_slide_numbers()
            .with_date_time();
        let handout_xml = handout.to_xml();
        assert!(!handout_xml.is_empty());
        write_summary_package(
            &format!("handout_bg_{name}.pptx"),
            &format!("Handout background: #{color}"),
            &format!(
                "Header/footer, slide numbers, and automatic date\nXML bytes: {}",
                handout_xml.len()
            ),
            4,
        )?;
        println!("  ✓ #{color} ({name})");
    }

    Ok(())
}

/// Test 3: combine a twelve-slide package summary with a six-up handout model.
fn test_multi_slide_handout() -> Result<(), Box<dyn StdError>> {
    println!("\nTest 3: multi-slide handout package round trip...");

    let handout = Master::new().with_layout(Layout::SixSlides);
    let handout_xml = handout.to_xml();
    write_summary_package(
        "handout_12_slides.pptx",
        "Twelve-slide handout",
        &format!(
            "Six slides per page; two expected handout pages\nXML bytes: {}",
            handout_xml.len()
        ),
        12,
    )?;
    println!("  ✓ Saved and reopened handout_12_slides.pptx");
    Ok(())
}

/// Test 4: combine handout, section, and timing owners in several package cases.
fn test_full_featured_handout() -> Result<(), Box<dyn StdError>> {
    println!("\nTest 4: full typed handout, section, and animation cases...");

    let sections = sections_xml()?;
    let animations = animations_xml()?;
    let handout = Master::new()
        .with_layout(Layout::ThreeSlides)
        .with_background_color("F5F5F5")
        .with_header("Confidential - Internal Use Only")
        .with_footer("© 2024 Example Corp")
        .with_slide_numbers()
        .with_date_time();
    let handout_xml = handout.to_xml();

    let full_body = format!(
        "Eight-slide summary deck\nSections XML: {} bytes\nAnimations XML: {} bytes\nHandout XML: {} bytes",
        sections.len(),
        animations.len(),
        handout_xml.len()
    );
    write_summary_package(
        "handout_full_featured.pptx",
        "Full-featured handout",
        &full_body,
        8,
    )?;
    println!("  ✓ Saved and reopened handout_full_featured.pptx");

    let fixed_date = Master::new()
        .with_header("Header Text Here")
        .with_footer("Footer Text Here")
        .with_slide_numbers()
        .with_fixed_date("December 5, 2024")
        .to_xml();
    write_summary_package(
        "handout_header_footer_test.pptx",
        "Header/footer handout",
        &format!("Fixed-date handout XML bytes: {}", fixed_date.len()),
        4,
    )?;

    let auto_date = Master::new()
        .with_layout(Layout::ThreeSlides)
        .with_background_color("F5F5F5")
        .with_header("Confidential - Internal Use Only")
        .with_footer("© 2024 Example Corp")
        .with_slide_numbers()
        .with_date_time()
        .to_xml();
    write_summary_package(
        "handout_auto_date_test.pptx",
        "Automatic-date handout",
        &format!("Automatic-date handout XML bytes: {}", auto_date.len()),
        4,
    )?;

    test_8slides_handout(&handout_xml)?;
    test_8slides_sections_handout(&sections, &handout_xml)?;
    test_8slides_animations_handout(&animations, &handout_xml)?;
    test_8slides_sections_animations_no_handout(&sections, &animations)?;
    Ok(())
}

fn animations_xml() -> Result<String, Box<dyn StdError>> {
    let mut sequence = Sequence::new();
    sequence.add(
        EffectInstance::new(3, Effect::FlyIn)
            .with_trigger(Trigger::OnClick)
            .with_duration_ms(500),
    );
    sequence.add(
        EffectInstance::new(3, Effect::Fade)
            .with_trigger(Trigger::AfterPrevious)
            .with_duration_ms(750),
    );
    let xml = sequence.to_xml();
    let parsed = Sequence::parse_timing_xml(&xml)?;
    assert_eq!(parsed.animations.len(), 2);
    Ok(xml)
}

fn sections_xml() -> Result<String, Box<dyn StdError>> {
    let mut sections = SectionList::new();
    sections
        .add_section(Section::new("Introduction", "section-introduction").with_slides([256, 257]));
    sections.add_section(Section::new("Main Content", "section-content").with_slides([258, 259]));
    sections.add_section(Section::new("Demonstrations", "section-demos").with_slides([260, 261]));
    sections.add_section(Section::new("Conclusion", "section-conclusion").with_slides([262, 263]));
    let xml = sections.to_xml()?;
    assert_eq!(sections.sections().len(), 4);
    Ok(xml)
}

fn test_8slides_handout(handout_xml: &str) -> Result<(), Box<dyn StdError>> {
    write_summary_package(
        "test_8slides_handout.pptx",
        "Eight slides and handout",
        &format!("Handout XML bytes: {}", handout_xml.len()),
        8,
    )
}

fn test_8slides_sections_handout(
    sections_xml: &str,
    handout_xml: &str,
) -> Result<(), Box<dyn StdError>> {
    write_summary_package(
        "test_8slides_sections_handout.pptx",
        "Eight slides, sections, and handout",
        &format!(
            "Sections XML bytes: {}\nHandout XML bytes: {}",
            sections_xml.len(),
            handout_xml.len()
        ),
        8,
    )
}

fn test_8slides_animations_handout(
    animations_xml: &str,
    handout_xml: &str,
) -> Result<(), Box<dyn StdError>> {
    write_summary_package(
        "test_8slides_animations_handout.pptx",
        "Eight slides, animations, and handout",
        &format!(
            "Animations XML bytes: {}\nHandout XML bytes: {}",
            animations_xml.len(),
            handout_xml.len()
        ),
        8,
    )
}

fn test_8slides_sections_animations_no_handout(
    sections_xml: &str,
    animations_xml: &str,
) -> Result<(), Box<dyn StdError>> {
    write_summary_package(
        "test_8slides_sections_animations_no_handout.pptx",
        "Eight slides, sections, and animations",
        &format!(
            "Sections XML bytes: {}\nAnimations XML bytes: {}\nNo handout owner",
            sections_xml.len(),
            animations_xml.len()
        ),
        8,
    )
}

fn write_summary_package(
    path: &str,
    title: &str,
    body: &str,
    slide_count: usize,
) -> Result<(), Box<dyn StdError>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
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
    Ok(())
}
