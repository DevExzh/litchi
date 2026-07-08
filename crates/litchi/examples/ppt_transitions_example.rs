//! Example demonstrating PPT slide transitions
//!
//! This example creates a PPT file showcasing various slide transition effects
//! including speed variations, directions, and advance modes.
//!
//! Run with: cargo run --example ppt_transitions_example

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::transition::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use litchi::ole::ppt::writer::SlideTiming;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating PPT file with transitions...");

    let mut writer = PptWriter::new_widescreen();
    writer.set_property("Title", "Transition Effects Showcase");
    writer.set_property("Author", "Litchi Transition Demo");

    // Slide 1: Classic Transitions
    println!("Creating slide 1: Classic Transitions");
    let slide1 = writer.add_slide()?;
    writer.add_textbox(
        slide1,
        100,
        200,
        600,
        100,
        "Classic Transitions\n\nFade → Dissolve → Wipe",
    )?;

    let transition1 = TransitionInfo::with_type(TransitionType::Fade)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_timing(slide1, timing_from_transition(&transition1))?;

    // Slide 2: Directional Transitions
    println!("Creating slide 2: Directional Transitions");
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 100, 150, 600, 100, "Directional Transitions")?;
    writer.add_textbox(slide2, 100, 280, 600, 100, "Push from Left →")?;

    let transition2 = TransitionInfo::with_type(TransitionType::Push)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::FromLeft)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_timing(slide2, timing_from_transition(&transition2))?;

    // Slide 3: Wipe Transitions
    println!("Creating slide 3: Wipe Variations");
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 100, 150, 600, 100, "Wipe Transitions")?;
    writer.add_textbox(slide3, 100, 280, 600, 100, "Wipe from Bottom ↑")?;

    let transition3 = TransitionInfo::with_type(TransitionType::Wipe)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::FromBottom)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_timing(slide3, timing_from_transition(&transition3))?;

    // Slide 4: Split Transitions
    println!("Creating slide 4: Split Transitions");
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 100, 150, 600, 100, "Split Transitions")?;
    writer.add_textbox(slide4, 100, 280, 600, 100, "Split Horizontal ↔")?;

    let transition4 = TransitionInfo::with_type(TransitionType::Split)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::Horizontal)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 5: Box Transition
    println!("Creating slide 5: Box Transition");
    let slide5 = writer.add_slide()?;
    writer.add_textbox(slide5, 100, 150, 600, 100, "Box Transition")?;
    writer.add_textbox(slide5, 100, 280, 600, 100, "Box In →")?;

    let transition5 = TransitionInfo::with_type(TransitionType::Box)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::In)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 6: Blinds Transition
    println!("Creating slide 6: Blinds Transition");
    let slide6 = writer.add_slide()?;
    writer.add_textbox(slide6, 100, 150, 600, 100, "Blinds Transition")?;
    writer.add_textbox(slide6, 100, 280, 600, 100, "Vertical Blinds ↓")?;

    let transition6 = TransitionInfo::with_type(TransitionType::Blinds)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::Vertical)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 7: Checkerboard Transition
    println!("Creating slide 7: Checkerboard Transition");
    let slide7 = writer.add_slide()?;
    writer.add_textbox(slide7, 100, 150, 600, 100, "Checkerboard")?;
    writer.add_textbox(slide7, 100, 280, 600, 100, "Checkerboard Across")?;

    let transition7 = TransitionInfo::with_type(TransitionType::Checkerboard)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::Horizontal)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 8: Zoom Transition
    println!("Creating slide 8: Zoom Transition");
    let slide8 = writer.add_slide()?;
    writer.add_textbox(slide8, 100, 150, 600, 100, "Zoom Transition")?;
    writer.add_textbox(slide8, 100, 280, 600, 100, "Zoom In Effect")?;

    let transition8 = TransitionInfo::with_type(TransitionType::Zoom)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::In)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 9: Automatic Advance
    println!("Creating slide 9: Automatic Advance");
    let slide9 = writer.add_slide()?;
    writer.add_textbox(slide9, 100, 150, 600, 100, "Automatic Advance")?;
    writer.add_textbox(slide9, 100, 280, 600, 100, "Auto-advances after 3 seconds")?;

    let transition9 = TransitionInfo::with_type(TransitionType::Dissolve)
        .with_speed(TransitionSpeed::Slow)
        .with_advance_mode(AdvanceMode::Automatic)
        .with_advance_time(3000); // 3 seconds

    // Slide 10: Both Modes
    println!("Creating slide 10: Click or Auto Advance");
    let slide10 = writer.add_slide()?;
    writer.add_textbox(slide10, 100, 150, 600, 100, "Click or Wait")?;
    writer.add_textbox(slide10, 100, 280, 600, 100, "Click OR wait 5 seconds")?;

    let transition10 = TransitionInfo::with_type(TransitionType::Random)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::Both)
        .with_advance_time(5000); // 5 seconds

    // Slide 11: Modern Transitions - Morph
    println!("Creating slide 11: Modern Transitions");
    let slide11 = writer.add_slide()?;
    writer.add_textbox(slide11, 100, 150, 600, 100, "Modern Transitions")?;
    writer.add_textbox(
        slide11,
        100,
        280,
        600,
        100,
        "Morph Effect (PowerPoint 2016+)",
    )?;

    let transition11 = TransitionInfo::with_type(TransitionType::Morph)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 12: Speed Comparison
    println!("Creating slide 12: Speed Comparison");
    let slide12 = writer.add_slide()?;
    writer.add_textbox(slide12, 100, 100, 600, 80, "Speed Variations")?;
    writer.add_textbox(
        slide12,
        100,
        200,
        600,
        200,
        "Transition Speeds:\n\n• Slow (2 seconds)\n• Medium (1 second)\n• Fast (0.5 seconds)",
    )?;

    let transition12 = TransitionInfo::with_type(TransitionType::Fade)
        .with_speed(TransitionSpeed::Slow)
        .with_advance_mode(AdvanceMode::OnClick);

    // Set transitions for remaining slides
    writer.set_slide_timing(slide4, timing_from_transition(&transition4))?;
    writer.set_slide_timing(slide5, timing_from_transition(&transition5))?;
    writer.set_slide_timing(slide6, timing_from_transition(&transition6))?;
    writer.set_slide_timing(slide7, timing_from_transition(&transition7))?;
    writer.set_slide_timing(slide8, timing_from_transition(&transition8))?;
    writer.set_slide_timing(slide9, timing_from_transition(&transition9))?;
    writer.set_slide_timing(slide10, timing_from_transition(&transition10))?;
    writer.set_slide_timing(slide11, timing_from_transition(&transition11))?;
    writer.set_slide_timing(slide12, timing_from_transition(&transition12))?;

    // Save the presentation
    println!("Saving to ppt_transitions_showcase.ppt...");
    writer.save("ppt_transitions_showcase.ppt")?;

    println!("✅ Transition showcase created successfully!");
    println!("   - File: ppt_transitions_showcase.ppt");
    println!("   - 12 slides demonstrating:");
    println!("     * Classic transitions (Fade, Dissolve, Wipe)");
    println!("     * Directional effects (Push, Wipe, Split)");
    println!("     * Pattern transitions (Blinds, Checkerboard, Box)");
    println!("     * Modern effects (Zoom, Morph)");
    println!("     * Advance modes (Click, Auto, Both)");
    println!("     * Speed variations (Slow, Medium, Fast)");
    println!("\n🎯 Open ppt_transitions_showcase.ppt in Microsoft PowerPoint");
    println!("   Press F5 to enter Slide Show mode and observe the transitions between slides!");

    Ok(())
}

fn timing_from_transition(transition: &TransitionInfo) -> SlideTiming {
    match transition.advance_mode {
        AdvanceMode::OnClick => SlideTiming::on_click_only(),
        AdvanceMode::Automatic => {
            SlideTiming::auto_advance(transition.advance_time_ms.unwrap_or_default())
                .with_click_advance(false)
        },
        AdvanceMode::Both => {
            SlideTiming::auto_advance(transition.advance_time_ms.unwrap_or_default())
        },
    }
}
